#include <windows.h>
#include <windowsx.h>
#include <d2d1.h>
#include <wincodec.h>
#include <dwrite.h>
#include <shellapi.h>
#include <commdlg.h>
#include <filesystem>
#include <vector>
#include <algorithm>
#include <cmath>
#include <string>

#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "windowscodecs.lib")
#pragma comment(lib, "dwrite.lib")

// =====================
// グローバル変数
// =====================
ID2D1Factory* g_d2dFactory = nullptr;
ID2D1HwndRenderTarget* g_renderTarget = nullptr;
ID2D1Bitmap* g_bitmap = nullptr;
ID2D1SolidColorBrush* g_placeholderBrush = nullptr;
HWND g_hwnd = nullptr;

IWICImagingFactory* g_wicFactory = nullptr;
IDWriteFactory* g_dwriteFactory = nullptr;
IDWriteTextFormat* g_placeholderFormat = nullptr;

UINT g_imageWidth = 0;
UINT g_imageHeight = 0;

float g_zoom = 1.0f;
bool g_fitToWindow = true;
bool g_isEdgeDragging = false;
POINT g_dragStartPoint{};
float g_dragStartZoom = 1.0f;

enum class SortMode
{
    NameAsc,
    NameDesc,
    TimeAsc,
    TimeDesc
};

struct ImageEntry
{
    std::filesystem::path path;
    std::filesystem::file_time_type writeTime;
};

std::vector<ImageEntry> g_imageList;
size_t g_currentIndex = 0;
SortMode g_sortMode = SortMode::NameAsc;
std::filesystem::path g_currentImagePath;
const float g_zoomMin = 0.05f;
const float g_zoomMax = 20.0f;
const float g_edgeDragMargin = 12.0f;

constexpr int kMenuOpen = 1001;
constexpr int kMenuNext = 1002;
constexpr int kMenuPrev = 1003;
constexpr int kMenuFit = 1004;
constexpr int kMenuZoomIn = 1005;
constexpr int kMenuZoomOut = 1006;
constexpr int kMenuSortNameAsc = 1101;
constexpr int kMenuSortNameDesc = 1102;
constexpr int kMenuSortTimeAsc = 1103;
constexpr int kMenuSortTimeDesc = 1104;

// =====================
// 前方宣言
// =====================
void Render(HWND hwnd);
bool InitDirect2D(HWND hwnd);
bool InitWIC();
bool InitDirectWrite();
bool LoadImageFromFile(const wchar_t* path);
void CleanupResources();
void DiscardRenderTarget();
void RefreshImageList(const std::filesystem::path& imagePath);
bool LoadImageByIndex(size_t index);
void SetFitToWindow(bool fit);
void AdjustZoom(float factor, const POINT& screenPoint);
bool ShowOpenImageDialog(HWND hwnd);
void UpdateWindowSizeToImage(HWND hwnd, float drawWidth, float drawHeight);
void UpdateFitZoomFromWindow(HWND hwnd);
void UpdateWindowToZoomedImage();

bool IsImageFile(const std::filesystem::path& path)
{
    if (!path.has_extension())
    {
        return false;
    }
    std::wstring ext = path.extension().wstring();
    std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
    return ext == L".png" || ext == L".jpg" || ext == L".jpeg" || ext == L".bmp"
        || ext == L".gif" || ext == L".tif" || ext == L".tiff" || ext == L".webp";
}

void SortImageList()
{
    auto compareNameAsc = [](const ImageEntry& a, const ImageEntry& b)
    {
        return a.path.filename().wstring() < b.path.filename().wstring();
    };
    auto compareNameDesc = [&](const ImageEntry& a, const ImageEntry& b)
    {
        return compareNameAsc(b, a);
    };
    auto compareTimeAsc = [](const ImageEntry& a, const ImageEntry& b)
    {
        if (a.writeTime == b.writeTime)
        {
            return a.path.filename().wstring() < b.path.filename().wstring();
        }
        return a.writeTime < b.writeTime;
    };
    auto compareTimeDesc = [&](const ImageEntry& a, const ImageEntry& b)
    {
        return compareTimeAsc(b, a);
    };

    switch (g_sortMode)
    {
    case SortMode::NameAsc:
        std::sort(g_imageList.begin(), g_imageList.end(), compareNameAsc);
        break;
    case SortMode::NameDesc:
        std::sort(g_imageList.begin(), g_imageList.end(), compareNameDesc);
        break;
    case SortMode::TimeAsc:
        std::sort(g_imageList.begin(), g_imageList.end(), compareTimeAsc);
        break;
    case SortMode::TimeDesc:
        std::sort(g_imageList.begin(), g_imageList.end(), compareTimeDesc);
        break;
    }
}

// =====================
// ウィンドウプロシージャ
// =====================
LRESULT CALLBACK WndProc(
    HWND hwnd,
    UINT msg,
    WPARAM wParam,
    LPARAM lParam
)
{
    switch (msg)
    {
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        BeginPaint(hwnd, &ps);
        Render(hwnd);
        EndPaint(hwnd, &ps);
        return 0;
    }

    case WM_SIZE:
    {
        if (g_renderTarget)
        {
            UINT w = LOWORD(lParam);
            UINT h = HIWORD(lParam);
            g_renderTarget->Resize(D2D1::SizeU(w, h));
        }
        if (g_fitToWindow)
        {
            UpdateFitZoomFromWindow(hwnd);
            InvalidateRect(hwnd, nullptr, TRUE);
        }
        return 0;
    }

    case WM_DROPFILES:
    {
        HDROP drop = reinterpret_cast<HDROP>(wParam);
        UINT fileCount = DragQueryFileW(drop, 0xFFFFFFFF, nullptr, 0);
        if (fileCount > 0)
        {
            // 複数ファイルは先頭のみ読み込む（必要なら後で拡張）
            UINT pathLength = DragQueryFileW(drop, 0, nullptr, 0);
            if (pathLength > 0)
            {
                std::wstring path(pathLength + 1, L'\0');
                DragQueryFileW(drop, 0, path.data(), pathLength + 1);
                path.resize(pathLength);
                if (LoadImageFromFile(path.c_str()))
                {
                    RefreshImageList(path);
                    SetFitToWindow(true);
                    InvalidateRect(hwnd, nullptr, TRUE);
                }
            }
        }
        DragFinish(drop);
        return 0;
    }

    case WM_CONTEXTMENU:
    {
        HMENU menu = CreatePopupMenu();
        AppendMenu(menu, MF_STRING, kMenuOpen, L"Open...");
        AppendMenu(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenu(menu, MF_STRING, kMenuPrev, L"Previous");
        AppendMenu(menu, MF_STRING, kMenuNext, L"Next");
        AppendMenu(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenu(menu, MF_STRING, kMenuZoomIn, L"Zoom In");
        AppendMenu(menu, MF_STRING, kMenuZoomOut, L"Zoom Out");
        AppendMenu(menu, MF_STRING, kMenuFit, L"Fit to Window");
        AppendMenu(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenu(menu, MF_STRING, kMenuSortNameAsc, L"Sort: Name (A-Z)");
        AppendMenu(menu, MF_STRING, kMenuSortNameDesc, L"Sort: Name (Z-A)");
        AppendMenu(menu, MF_STRING, kMenuSortTimeAsc, L"Sort: Modified (Old-New)");
        AppendMenu(menu, MF_STRING, kMenuSortTimeDesc, L"Sort: Modified (New-Old)");

        if (g_imageList.size() < 2)
        {
            EnableMenuItem(menu, kMenuPrev, MF_BYCOMMAND | MF_GRAYED);
            EnableMenuItem(menu, kMenuNext, MF_BYCOMMAND | MF_GRAYED);
        }

        CheckMenuItem(menu, kMenuSortNameAsc, MF_BYCOMMAND | (g_sortMode == SortMode::NameAsc ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(menu, kMenuSortNameDesc, MF_BYCOMMAND | (g_sortMode == SortMode::NameDesc ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(menu, kMenuSortTimeAsc, MF_BYCOMMAND | (g_sortMode == SortMode::TimeAsc ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(menu, kMenuSortTimeDesc, MF_BYCOMMAND | (g_sortMode == SortMode::TimeDesc ? MF_CHECKED : MF_UNCHECKED));

        POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        TrackPopupMenu(menu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, nullptr);
        DestroyMenu(menu);
        return 0;
    }

    case WM_NCHITTEST:
    {
        LRESULT hit = DefWindowProc(hwnd, msg, wParam, lParam);
        if (hit == HTCLIENT)
        {
            return HTCAPTION;
        }
        return hit;
    }

    case WM_COMMAND:
    {
        switch (LOWORD(wParam))
        {
        case kMenuOpen:
            if (ShowOpenImageDialog(hwnd))
            {
                SetFitToWindow(true);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            return 0;
        case kMenuPrev:
            if (!g_imageList.empty())
            {
                size_t index = (g_currentIndex + g_imageList.size() - 1) % g_imageList.size();
                if (LoadImageByIndex(index))
                {
                    InvalidateRect(hwnd, nullptr, TRUE);
                }
            }
            return 0;
        case kMenuNext:
            if (!g_imageList.empty())
            {
                size_t index = (g_currentIndex + 1) % g_imageList.size();
                if (LoadImageByIndex(index))
                {
                    InvalidateRect(hwnd, nullptr, TRUE);
                }
            }
            return 0;
        case kMenuZoomIn:
        {
            POINT pt{};
            GetCursorPos(&pt);
            AdjustZoom(1.1f, pt);
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        }
        case kMenuZoomOut:
        {
            POINT pt{};
            GetCursorPos(&pt);
            AdjustZoom(1.0f / 1.1f, pt);
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        }
        case kMenuFit:
            SetFitToWindow(true);
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        case kMenuSortNameAsc:
            g_sortMode = SortMode::NameAsc;
            if (!g_currentImagePath.empty())
            {
                RefreshImageList(g_currentImagePath);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            return 0;
        case kMenuSortNameDesc:
            g_sortMode = SortMode::NameDesc;
            if (!g_currentImagePath.empty())
            {
                RefreshImageList(g_currentImagePath);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            return 0;
        case kMenuSortTimeAsc:
            g_sortMode = SortMode::TimeAsc;
            if (!g_currentImagePath.empty())
            {
                RefreshImageList(g_currentImagePath);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            return 0;
        case kMenuSortTimeDesc:
            g_sortMode = SortMode::TimeDesc;
            if (!g_currentImagePath.empty())
            {
                RefreshImageList(g_currentImagePath);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            return 0;
        }
        break;
    }

    case WM_MOUSEWHEEL:
    {
        if (!g_bitmap)
        {
            return 0;
        }
        int delta = GET_WHEEL_DELTA_WPARAM(wParam);
        float steps = static_cast<float>(delta) / WHEEL_DELTA;
        float factor = std::pow(1.1f, steps);
        POINT pt{};
        GetCursorPos(&pt);
        AdjustZoom(factor, pt);
        InvalidateRect(hwnd, nullptr, TRUE);
        return 0;
    }

    case WM_LBUTTONDOWN:
    {
        if (!g_bitmap)
        {
            return 0;
        }
        POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        RECT rc{};
        GetClientRect(hwnd, &rc);
        bool nearEdge = pt.x <= g_edgeDragMargin || pt.y <= g_edgeDragMargin
            || pt.x >= (rc.right - g_edgeDragMargin) || pt.y >= (rc.bottom - g_edgeDragMargin);
        if (nearEdge && g_renderTarget && g_imageWidth > 0 && g_imageHeight > 0)
        {
            g_fitToWindow = false;
            g_isEdgeDragging = true;
            g_dragStartPoint = pt;
            g_dragStartZoom = g_zoom;
            UpdateWindowToZoomedImage();
            SetCapture(hwnd);
        }
        return 0;
    }

    case WM_MOUSEMOVE:
    {
        if (g_isEdgeDragging && (wParam & MK_LBUTTON))
        {
            POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
            float delta = static_cast<float>(pt.y - g_dragStartPoint.y) * 0.005f;
            float zoom = g_dragStartZoom * (1.0f + delta);
            g_zoom = max(g_zoomMin, min(zoom, g_zoomMax));
            UpdateWindowToZoomedImage();
            InvalidateRect(hwnd, nullptr, TRUE);
        }
        return 0;
    }

    case WM_LBUTTONUP:
    {
        if (g_isEdgeDragging)
        {
            g_isEdgeDragging = false;
            ReleaseCapture();
        }
        return 0;
    }

    case WM_KEYDOWN:
    {
        if (wParam == VK_RIGHT && !g_imageList.empty())
        {
            size_t index = (g_currentIndex + 1) % g_imageList.size();
            if (LoadImageByIndex(index))
            {
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            return 0;
        }
        if (wParam == VK_LEFT && !g_imageList.empty())
        {
            size_t index = (g_currentIndex + g_imageList.size() - 1) % g_imageList.size();
            if (LoadImageByIndex(index))
            {
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            return 0;
        }
        if (wParam == VK_ADD || wParam == VK_OEM_PLUS)
        {
            POINT pt{};
            GetCursorPos(&pt);
            AdjustZoom(1.1f, pt);
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        }
        if (wParam == VK_SUBTRACT || wParam == VK_OEM_MINUS)
        {
            POINT pt{};
            GetCursorPos(&pt);
            AdjustZoom(1.0f / 1.1f, pt);
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        }
        if ((GetKeyState(VK_CONTROL) & 0x8000) && wParam == '0')
        {
            SetFitToWindow(true);
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        }
        if ((GetKeyState(VK_CONTROL) & 0x8000) && wParam == '1')
        {
            g_fitToWindow = false;
            g_zoom = 1.0f;
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        }
        return 0;
    }

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProc(hwnd, msg, wParam, lParam);
}

// =====================
// Direct2D 初期化
// =====================
bool InitDirect2D(HWND hwnd)
{
    if (FAILED(D2D1CreateFactory(
        D2D1_FACTORY_TYPE_SINGLE_THREADED,
        &g_d2dFactory)))
    {
        return false;
    }

    RECT rc;
    GetClientRect(hwnd, &rc);

    if (FAILED(g_d2dFactory->CreateHwndRenderTarget(
        D2D1::RenderTargetProperties(),
        D2D1::HwndRenderTargetProperties(
            hwnd,
            D2D1::SizeU(
                rc.right - rc.left,
                rc.bottom - rc.top
            )
        ),
        &g_renderTarget)))
    {
        return false;
    }

    if (FAILED(g_renderTarget->CreateSolidColorBrush(
        D2D1::ColorF(D2D1::ColorF::White),
        &g_placeholderBrush)))
    {
        return false;
    }

    return true;
}

void DiscardRenderTarget()
{
    if (g_placeholderBrush)
    {
        g_placeholderBrush->Release();
        g_placeholderBrush = nullptr;
    }
    if (g_bitmap)
    {
        g_bitmap->Release();
        g_bitmap = nullptr;
    }
    if (g_renderTarget)
    {
        g_renderTarget->Release();
        g_renderTarget = nullptr;
    }
    g_imageWidth = 0;
    g_imageHeight = 0;
}

void CleanupResources()
{
    DiscardRenderTarget();

    if (g_placeholderFormat)
    {
        g_placeholderFormat->Release();
        g_placeholderFormat = nullptr;
    }
    if (g_dwriteFactory)
    {
        g_dwriteFactory->Release();
        g_dwriteFactory = nullptr;
    }
    if (g_wicFactory)
    {
        g_wicFactory->Release();
        g_wicFactory = nullptr;
    }
    if (g_d2dFactory)
    {
        g_d2dFactory->Release();
        g_d2dFactory = nullptr;
    }
}

// =====================
// WIC 初期化
// =====================
bool InitWIC()
{
    HRESULT hr = CoCreateInstance(
        CLSID_WICImagingFactory,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&g_wicFactory)
    );
    return SUCCEEDED(hr);
}

// =====================
// DirectWrite 初期化
// =====================
bool InitDirectWrite()
{
    HRESULT hr = DWriteCreateFactory(
        DWRITE_FACTORY_TYPE_SHARED,
        __uuidof(IDWriteFactory),
        reinterpret_cast<IUnknown**>(&g_dwriteFactory)
    );
    if (FAILED(hr))
    {
        return false;
    }

    hr = g_dwriteFactory->CreateTextFormat(
        L"Segoe UI",
        nullptr,
        DWRITE_FONT_WEIGHT_REGULAR,
        DWRITE_FONT_STYLE_NORMAL,
        DWRITE_FONT_STRETCH_NORMAL,
        24.0f,
        L"ja-jp",
        &g_placeholderFormat
    );
    if (FAILED(hr))
    {
        return false;
    }

    g_placeholderFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
    g_placeholderFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

    return true;
}

// =====================
// 画像ロード
// =====================
bool LoadImageFromFile(const wchar_t* path)
{
    IWICBitmapDecoder* decoder = nullptr;
    IWICBitmapFrameDecode* frame = nullptr;
    IWICFormatConverter* converter = nullptr;

    if (g_bitmap)
    {
        g_bitmap->Release();
        g_bitmap = nullptr;
    }
    g_imageWidth = 0;
    g_imageHeight = 0;

    HRESULT hr = g_wicFactory->CreateDecoderFromFilename(
        path,
        nullptr,
        GENERIC_READ,
        WICDecodeMetadataCacheOnDemand,
        &decoder
    );
    if (FAILED(hr)) goto cleanup;

    hr = decoder->GetFrame(0, &frame);
    if (FAILED(hr)) goto cleanup;

    hr = frame->GetSize(&g_imageWidth, &g_imageHeight);
    if (FAILED(hr)) goto cleanup;

    hr = g_wicFactory->CreateFormatConverter(&converter);
    if (FAILED(hr)) goto cleanup;

    hr = converter->Initialize(
        frame,
        GUID_WICPixelFormat32bppPBGRA,
        WICBitmapDitherTypeNone,
        nullptr,
        0.0,
        WICBitmapPaletteTypeCustom
    );
    if (FAILED(hr)) goto cleanup;

    hr = g_renderTarget->CreateBitmapFromWicBitmap(
        converter,
        nullptr,
        &g_bitmap
    );

cleanup:
    if (decoder) decoder->Release();
    if (frame) frame->Release();
    if (converter) converter->Release();

    return SUCCEEDED(hr);
}

void RefreshImageList(const std::filesystem::path& imagePath)
{
    g_imageList.clear();
    g_currentImagePath = imagePath;
    g_currentIndex = 0;

    std::filesystem::path dir = imagePath.parent_path();
    if (dir.empty())
    {
        dir = std::filesystem::current_path();
    }

    std::error_code ec;
    for (const auto& entry : std::filesystem::directory_iterator(dir, ec))
    {
        if (ec)
        {
            break;
        }
        if (!entry.is_regular_file(ec))
        {
            continue;
        }
        std::filesystem::path filePath = entry.path();
        if (!IsImageFile(filePath))
        {
            continue;
        }
        std::filesystem::file_time_type time = entry.last_write_time(ec);
        g_imageList.push_back({ filePath, time });
    }

    SortImageList();

    for (size_t i = 0; i < g_imageList.size(); ++i)
    {
        if (g_imageList[i].path == g_currentImagePath)
        {
            g_currentIndex = i;
            break;
        }
    }
}

bool LoadImageByIndex(size_t index)
{
    if (g_imageList.empty() || index >= g_imageList.size())
    {
        return false;
    }
    g_currentIndex = index;
    g_currentImagePath = g_imageList[index].path;
    bool result = LoadImageFromFile(g_currentImagePath.c_str());
    if (result)
    {
        SetFitToWindow(true);
    }
    return result;
}

void SetFitToWindow(bool fit)
{
    g_fitToWindow = fit;
    if (fit)
    {
        UpdateFitZoomFromWindow(nullptr);
        UpdateWindowToZoomedImage();
    }
}

void AdjustZoom(float factor, const POINT& screenPoint)
{
    (void)screenPoint;
    float newScale = g_zoom * factor;
    g_zoom = max(g_zoomMin, min(newScale, g_zoomMax));
    g_fitToWindow = false;
    UpdateWindowToZoomedImage();
}

bool ShowOpenImageDialog(HWND hwnd)
{
    wchar_t filePath[MAX_PATH] = L"";
    OPENFILENAME ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = filePath;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = L"Image Files\0*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.webp\0All Files\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;

    if (!GetOpenFileName(&ofn))
    {
        return false;
    }

    if (LoadImageFromFile(filePath))
    {
        RefreshImageList(filePath);
        return true;
    }
    return false;
}

void UpdateWindowSizeToImage(HWND hwnd, float drawWidth, float drawHeight)
{
    if (!hwnd || drawWidth <= 0.0f || drawHeight <= 0.0f)
    {
        return;
    }

    RECT windowRect{};
    RECT clientRect{};
    GetWindowRect(hwnd, &windowRect);
    GetClientRect(hwnd, &clientRect);

    int currentClientWidth = clientRect.right - clientRect.left;
    int currentClientHeight = clientRect.bottom - clientRect.top;
    int targetClientWidth = static_cast<int>(std::lround(drawWidth));
    int targetClientHeight = static_cast<int>(std::lround(drawHeight));

    if (abs(targetClientWidth - currentClientWidth) < 2 && abs(targetClientHeight - currentClientHeight) < 2)
    {
        return;
    }

    int windowWidth = windowRect.right - windowRect.left;
    int windowHeight = windowRect.bottom - windowRect.top;
    int frameWidth = windowWidth - currentClientWidth;
    int frameHeight = windowHeight - currentClientHeight;

    int targetWindowWidth = targetClientWidth + frameWidth;
    int targetWindowHeight = targetClientHeight + frameHeight;

    SetWindowPos(
        hwnd,
        nullptr,
        0,
        0,
        targetWindowWidth,
        targetWindowHeight,
        SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE
    );
}

void UpdateFitZoomFromWindow(HWND hwnd)
{
    (void)hwnd;
    if (!g_renderTarget || g_imageWidth == 0 || g_imageHeight == 0)
    {
        g_zoom = 1.0f;
        return;
    }

    D2D1_SIZE_F rtSize = g_renderTarget->GetSize();
    float scaleX = rtSize.width / static_cast<float>(g_imageWidth);
    float scaleY = rtSize.height / static_cast<float>(g_imageHeight);
    g_zoom = max(g_zoomMin, min(min(scaleX, scaleY), g_zoomMax));
}

void UpdateWindowToZoomedImage()
{
    if (!g_hwnd || g_imageWidth == 0 || g_imageHeight == 0)
    {
        return;
    }
    UpdateWindowSizeToImage(g_hwnd, g_imageWidth * g_zoom, g_imageHeight * g_zoom);
}

// =====================
// 描画
// =====================
void Render(HWND hwnd)
{
    if (!g_renderTarget)
    {
        if (!InitDirect2D(hwnd))
        {
            return;
        }
    }

    g_renderTarget->BeginDraw();

    if (g_bitmap)
    {
        float scale = g_zoom;
        float drawWidth = g_imageWidth * scale;
        float drawHeight = g_imageHeight * scale;

        UpdateWindowSizeToImage(hwnd, drawWidth, drawHeight);

        if (g_renderTarget && drawWidth > 0.0f && drawHeight > 0.0f)
        {
            D2D1_SIZE_F rtSize = g_renderTarget->GetSize();
            UINT targetWidth = static_cast<UINT>(max(1.0f, std::lround(drawWidth)));
            UINT targetHeight = static_cast<UINT>(max(1.0f, std::lround(drawHeight)));
            if (rtSize.width != targetWidth || rtSize.height != targetHeight)
            {
                g_renderTarget->Resize(D2D1::SizeU(targetWidth, targetHeight));
            }
        }

        g_renderTarget->Clear(D2D1::ColorF(D2D1::ColorF::Black));

        D2D1_RECT_F dest = D2D1::RectF(
            0.0f,
            0.0f,
            drawWidth,
            drawHeight
        );

        g_renderTarget->DrawBitmap(
            g_bitmap,
            dest,
            1.0f,
            D2D1_BITMAP_INTERPOLATION_MODE_LINEAR
        );
    }
    else
    {
        g_renderTarget->Clear(D2D1::ColorF(0.12f, 0.12f, 0.12f));

        const wchar_t* placeholderText = L"Drop image here";
        D2D1_SIZE_F rtSize = g_renderTarget->GetSize();
        D2D1_RECT_F textRect = D2D1::RectF(
            0.0f,
            0.0f,
            rtSize.width,
            rtSize.height
        );

        if (g_placeholderBrush && g_placeholderFormat)
        {
            g_renderTarget->DrawTextW(
                placeholderText,
                static_cast<UINT32>(wcslen(placeholderText)),
                g_placeholderFormat,
                textRect,
                g_placeholderBrush
            );
        }
    }

    HRESULT hr = g_renderTarget->EndDraw();
    if (hr == D2DERR_RECREATE_TARGET)
    {
        DiscardRenderTarget();
        InvalidateRect(hwnd, nullptr, TRUE);
    }
}

// =====================
// エントリポイント
// =====================
int WINAPI wWinMain(
    HINSTANCE hInstance,
    HINSTANCE,
    PWSTR,
    int nCmdShow
)
{
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    const wchar_t CLASS_NAME[] = L"FloatVisionWindow";

    WNDCLASS wc{};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;

    RegisterClass(&wc);

    HWND hwnd = CreateWindowEx(
        0,
        CLASS_NAME,
        L"FloatVision",
        WS_POPUP,
        CW_USEDEFAULT, CW_USEDEFAULT,
        800, 600,
        nullptr,
        nullptr,
        hInstance,
        nullptr
    );
    g_hwnd = hwnd;

    if (!hwnd)
    {
        MessageBox(nullptr, L"CreateWindowEx failed", L"Error", MB_OK);
        return 0;
    }

    DragAcceptFiles(hwnd, TRUE);

    if (!InitWIC())
    {
        MessageBox(nullptr, L"WIC init failed", L"Error", MB_OK);
        return 0;
    }

    if (!InitDirectWrite())
    {
        MessageBox(nullptr, L"DirectWrite init failed", L"Error", MB_OK);
        return 0;
    }

    if (!InitDirect2D(hwnd))
    {
        MessageBox(hwnd, L"Direct2D init failed", L"Error", MB_OK);
        return 0;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    int argc = 0;
    wchar_t** argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    bool loadedImage = false;
    if (argv && argc > 1)
    {
        loadedImage = LoadImageFromFile(argv[1]);
        if (loadedImage)
        {
            RefreshImageList(argv[1]);
            SetFitToWindow(true);
            UpdateWindowSizeToImage(hwnd, g_imageWidth * g_zoom, g_imageHeight * g_zoom);
        }
    }
    if (argv)
    {
        LocalFree(argv);
    }
    if (!loadedImage)
    {
        // 仮画像表示（g_bitmap が未設定のため Render でプレースホルダ表示）
    }
    InvalidateRect(hwnd, nullptr, TRUE);

    MSG msg{};
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    CleanupResources();
    CoUninitialize();
    return 0;
}
