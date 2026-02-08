#define NOMINMAX
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
#include <cstdlib>
#include <string>
#include "resource.h"

#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "windowscodecs.lib")
#pragma comment(lib, "dwrite.lib")

#ifndef FLOATVISION_GLOBALS_DEFINED
#define FLOATVISION_GLOBALS_DEFINED
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
IWICBitmapSource* g_wicSource = nullptr;

UINT g_imageWidth = 0;
UINT g_imageHeight = 0;
bool g_imageHasAlpha = false;

float g_zoom = 1.0f;
bool g_fitToWindow = true;
bool g_isEdgeDragging = false;
POINT g_dragStartPoint{};
float g_dragStartZoom = 1.0f;
float g_dragStartScale = 1.0f;

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
bool g_alwaysOnTop = false;
std::wstring g_iniPath;
POINT g_windowPos{ CW_USEDEFAULT, CW_USEDEFAULT };
bool g_hasSavedWindowPos = false;

constexpr int kMenuOpen = 1001;
constexpr int kMenuNext = 1002;
constexpr int kMenuPrev = 1003;
constexpr int kMenuZoomIn = 1005;
constexpr int kMenuZoomOut = 1006;
constexpr int kMenuOriginalSize = 1007;
constexpr int kMenuAlwaysOnTop = 1008;
constexpr int kMenuExit = 1009;
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
void NavigateImage(int delta);
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
void UpdateZoomToFitScreen(HWND hwnd);
void LoadSettings();
void SaveSettings();
void ApplyAlwaysOnTop();
void LoadWindowPlacement();
void SaveWindowPlacement();
void UpdateLayeredStyle(bool enable);
bool UpdateLayeredWindowFromWic(HWND hwnd, float drawWidth, float drawHeight);
bool QueryPixelFormatHasAlpha(const WICPixelFormatGUID& format);
#endif

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
                    UpdateZoomToFitScreen(hwnd);
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
        AppendMenu(menu, MF_STRING, kMenuOriginalSize, L"Original Size");
        AppendMenu(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenu(menu, MF_STRING, kMenuAlwaysOnTop, L"Always on Top");
        AppendMenu(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenu(menu, MF_STRING, kMenuSortNameAsc, L"Sort: Name (A-Z)");
        AppendMenu(menu, MF_STRING, kMenuSortNameDesc, L"Sort: Name (Z-A)");
        AppendMenu(menu, MF_STRING, kMenuSortTimeAsc, L"Sort: Modified (Old-New)");
        AppendMenu(menu, MF_STRING, kMenuSortTimeDesc, L"Sort: Modified (New-Old)");
        AppendMenu(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenu(menu, MF_STRING, kMenuExit, L"Exit");

        if (g_imageList.size() < 2)
        {
            EnableMenuItem(menu, kMenuPrev, MF_BYCOMMAND | MF_GRAYED);
            EnableMenuItem(menu, kMenuNext, MF_BYCOMMAND | MF_GRAYED);
        }

        CheckMenuItem(menu, kMenuSortNameAsc, MF_BYCOMMAND | (g_sortMode == SortMode::NameAsc ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(menu, kMenuSortNameDesc, MF_BYCOMMAND | (g_sortMode == SortMode::NameDesc ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(menu, kMenuSortTimeAsc, MF_BYCOMMAND | (g_sortMode == SortMode::TimeAsc ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(menu, kMenuSortTimeDesc, MF_BYCOMMAND | (g_sortMode == SortMode::TimeDesc ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(menu, kMenuAlwaysOnTop, MF_BYCOMMAND | (g_alwaysOnTop ? MF_CHECKED : MF_UNCHECKED));

        POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        TrackPopupMenu(menu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, nullptr);
        DestroyMenu(menu);
        return 0;
    }

    case WM_COMMAND:
    {
        switch (LOWORD(wParam))
        {
        case kMenuOpen:
            if (ShowOpenImageDialog(hwnd))
            {
                UpdateZoomToFitScreen(hwnd);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            return 0;
        case kMenuPrev:
            NavigateImage(-1);
            return 0;
        case kMenuNext:
            NavigateImage(1);
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
        case kMenuOriginalSize:
            g_fitToWindow = false;
            g_zoom = 1.0f;
            UpdateWindowToZoomedImage();
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        case kMenuAlwaysOnTop:
            g_alwaysOnTop = !g_alwaysOnTop;
            ApplyAlwaysOnTop();
            SaveSettings();
            return 0;
        case kMenuExit:
            DestroyWindow(hwnd);
            return 0;
        case kMenuSortNameAsc:
            g_sortMode = SortMode::NameAsc;
            if (!g_currentImagePath.empty())
            {
                RefreshImageList(g_currentImagePath);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            SaveSettings();
            return 0;
        case kMenuSortNameDesc:
            g_sortMode = SortMode::NameDesc;
            if (!g_currentImagePath.empty())
            {
                RefreshImageList(g_currentImagePath);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            SaveSettings();
            return 0;
        case kMenuSortTimeAsc:
            g_sortMode = SortMode::TimeAsc;
            if (!g_currentImagePath.empty())
            {
                RefreshImageList(g_currentImagePath);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            SaveSettings();
            return 0;
        case kMenuSortTimeDesc:
            g_sortMode = SortMode::TimeDesc;
            if (!g_currentImagePath.empty())
            {
                RefreshImageList(g_currentImagePath);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            SaveSettings();
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

    case WM_XBUTTONUP:
    {
        if (GET_XBUTTON_WPARAM(wParam) == XBUTTON1)
        {
            NavigateImage(-1);
            return 0;
        }
        if (GET_XBUTTON_WPARAM(wParam) == XBUTTON2)
        {
            NavigateImage(1);
            return 0;
        }
        break;
    }

    case WM_LBUTTONDOWN:
    {
        POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        RECT rc{};
        GetClientRect(hwnd, &rc);
        bool nearEdge = pt.x <= g_edgeDragMargin || pt.y <= g_edgeDragMargin
            || pt.x >= (rc.right - g_edgeDragMargin) || pt.y >= (rc.bottom - g_edgeDragMargin);
        if (g_bitmap && nearEdge && g_renderTarget && g_imageWidth > 0 && g_imageHeight > 0)
        {
            g_fitToWindow = false;
            g_isEdgeDragging = true;
            g_dragStartPoint = pt;
            g_dragStartZoom = g_zoom;
            g_dragStartScale = std::max(1.0f, (std::min)(static_cast<float>(rc.right - rc.left), static_cast<float>(rc.bottom - rc.top)));
            UpdateWindowToZoomedImage();
            SetCapture(hwnd);
            return 0;
        }
        ReleaseCapture();
        SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, lParam);
        return 0;
    }

    case WM_MOUSEMOVE:
    {
        if (g_isEdgeDragging && (wParam & MK_LBUTTON))
        {
            POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
            float delta = static_cast<float>(pt.y - g_dragStartPoint.y);
            float nextScale = std::max(1.0f, g_dragStartScale + delta);
            float zoom = (g_dragStartScale > 0.0f) ? (g_dragStartZoom * (nextScale / g_dragStartScale)) : g_dragStartZoom;
            g_zoom = std::max(g_zoomMin, (std::min)(zoom, g_zoomMax));
            UpdateWindowToZoomedImage();
            InvalidateRect(hwnd, nullptr, TRUE);
        }
        else
        {
            POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
            RECT rc{};
            GetClientRect(hwnd, &rc);
            bool nearEdge = pt.x <= g_edgeDragMargin || pt.y <= g_edgeDragMargin
                || pt.x >= (rc.right - g_edgeDragMargin) || pt.y >= (rc.bottom - g_edgeDragMargin);
            SetCursor(LoadCursor(nullptr, nearEdge ? IDC_SIZENWSE : IDC_ARROW));
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
            NavigateImage(1);
            return 0;
        }
        if (wParam == VK_LEFT && !g_imageList.empty())
        {
            NavigateImage(-1);
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
        SaveWindowPlacement();
        SaveSettings();
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

    g_renderTarget->SetDpi(96.0f, 96.0f);

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
    if (g_wicSource)
    {
        g_wicSource->Release();
        g_wicSource = nullptr;
    }
    g_imageWidth = 0;
    g_imageHeight = 0;
    g_imageHasAlpha = false;
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
    WICPixelFormatGUID pixelFormat = GUID_WICPixelFormatDontCare;

    if (g_bitmap)
    {
        g_bitmap->Release();
        g_bitmap = nullptr;
    }
    if (g_wicSource)
    {
        g_wicSource->Release();
        g_wicSource = nullptr;
    }
    g_imageWidth = 0;
    g_imageHeight = 0;
    g_imageHasAlpha = false;

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

    hr = frame->GetPixelFormat(&pixelFormat);
    if (SUCCEEDED(hr))
    {
        g_imageHasAlpha = QueryPixelFormatHasAlpha(pixelFormat);
    }

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
    if (SUCCEEDED(hr))
    {
        g_wicSource = converter;
        g_wicSource->AddRef();
    }

cleanup:
    if (decoder) decoder->Release();
    if (frame) frame->Release();
    if (converter) converter->Release();

    UpdateLayeredStyle(g_imageHasAlpha);
    return SUCCEEDED(hr);
}

bool QueryPixelFormatHasAlpha(const WICPixelFormatGUID& format)
{
    if (!g_wicFactory)
    {
        return false;
    }

    IWICComponentInfo* componentInfo = nullptr;
    IWICPixelFormatInfo2* formatInfo = nullptr;
    bool hasAlpha = false;

    HRESULT hr = g_wicFactory->CreateComponentInfo(format, &componentInfo);
    if (SUCCEEDED(hr))
    {
        hr = componentInfo->QueryInterface(IID_PPV_ARGS(&formatInfo));
    }
    if (SUCCEEDED(hr))
    {
        BOOL supportsTransparency = FALSE;
        if (SUCCEEDED(formatInfo->SupportsTransparency(&supportsTransparency)))
        {
            hasAlpha = supportsTransparency == TRUE;
        }
    }

    if (formatInfo) formatInfo->Release();
    if (componentInfo) componentInfo->Release();

    return hasAlpha;
}

void NavigateImage(int delta)
{
    if (g_imageList.empty())
    {
        return;
    }

    size_t count = g_imageList.size();
    size_t index = (g_currentIndex + count + (delta % static_cast<int>(count))) % count;
    if (LoadImageByIndex(index) && g_hwnd)
    {
        InvalidateRect(g_hwnd, nullptr, TRUE);
    }
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
        UpdateZoomToFitScreen(g_hwnd);
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
    g_zoom = std::max(g_zoomMin, (std::min)(newScale, g_zoomMax));
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
        UpdateZoomToFitScreen(hwnd);
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
    g_zoom = std::max(g_zoomMin, (std::min)((std::min)(scaleX, scaleY), g_zoomMax));
}

void UpdateWindowToZoomedImage()
{
    if (!g_hwnd || g_imageWidth == 0 || g_imageHeight == 0)
    {
        return;
    }
    UpdateWindowSizeToImage(g_hwnd, g_imageWidth * g_zoom, g_imageHeight * g_zoom);
}

void UpdateZoomToFitScreen(HWND hwnd)
{
    if (g_imageWidth == 0 || g_imageHeight == 0)
    {
        g_zoom = 1.0f;
        return;
    }

    HMONITOR monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    MONITORINFO info{};
    info.cbSize = sizeof(info);
    if (monitor && GetMonitorInfo(monitor, &info))
    {
        int workWidth = info.rcWork.right - info.rcWork.left;
        int workHeight = info.rcWork.bottom - info.rcWork.top;
        if (workWidth > 0 && workHeight > 0)
        {
            float scaleX = static_cast<float>(workWidth) / static_cast<float>(g_imageWidth);
            float scaleY = static_cast<float>(workHeight) / static_cast<float>(g_imageHeight);
            g_zoom = (std::min)(1.0f, (std::min)(scaleX, scaleY));
            g_zoom = std::max(g_zoomMin, (std::min)(g_zoom, g_zoomMax));
            g_fitToWindow = (g_zoom < 1.0f);
            UpdateWindowToZoomedImage();
            return;
        }
    }

    g_zoom = 1.0f;
    g_fitToWindow = false;
    UpdateWindowToZoomedImage();
}

void LoadSettings()
{
    if (g_iniPath.empty())
    {
        return;
    }

    wchar_t buffer[32]{};
    GetPrivateProfileStringW(L"Settings", L"SortMode", L"0", buffer, 32, g_iniPath.c_str());
    int sortValue = _wtoi(buffer);
    if (sortValue < 0 || sortValue > 3)
    {
        sortValue = 0;
    }
    g_sortMode = static_cast<SortMode>(sortValue);

    GetPrivateProfileStringW(L"Settings", L"AlwaysOnTop", L"0", buffer, 32, g_iniPath.c_str());
    g_alwaysOnTop = (_wtoi(buffer) != 0);
}

void SaveSettings()
{
    if (g_iniPath.empty())
    {
        return;
    }

    wchar_t buffer[32]{};
    _snwprintf_s(buffer, _TRUNCATE, L"%d", static_cast<int>(g_sortMode));
    WritePrivateProfileStringW(L"Settings", L"SortMode", buffer, g_iniPath.c_str());

    _snwprintf_s(buffer, _TRUNCATE, L"%d", g_alwaysOnTop ? 1 : 0);
    WritePrivateProfileStringW(L"Settings", L"AlwaysOnTop", buffer, g_iniPath.c_str());
}

void ApplyAlwaysOnTop()
{
    if (!g_hwnd)
    {
        return;
    }
    SetWindowPos(
        g_hwnd,
        g_alwaysOnTop ? HWND_TOPMOST : HWND_NOTOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE
    );
}

void LoadWindowPlacement()
{
    if (g_iniPath.empty())
    {
        return;
    }

    wchar_t buffer[32]{};
    GetPrivateProfileStringW(L"Window", L"X", L"", buffer, 32, g_iniPath.c_str());
    if (buffer[0] != L'\0')
    {
        g_windowPos.x = _wtoi(buffer);
        g_hasSavedWindowPos = true;
    }
    GetPrivateProfileStringW(L"Window", L"Y", L"", buffer, 32, g_iniPath.c_str());
    if (buffer[0] != L'\0')
    {
        g_windowPos.y = _wtoi(buffer);
        g_hasSavedWindowPos = true;
    }
}

void SaveWindowPlacement()
{
    if (!g_hwnd || g_iniPath.empty())
    {
        return;
    }

    RECT rect{};
    if (!GetWindowRect(g_hwnd, &rect))
    {
        return;
    }

    wchar_t buffer[32]{};
    _snwprintf_s(buffer, _TRUNCATE, L"%d", rect.left);
    WritePrivateProfileStringW(L"Window", L"X", buffer, g_iniPath.c_str());
    _snwprintf_s(buffer, _TRUNCATE, L"%d", rect.top);
    WritePrivateProfileStringW(L"Window", L"Y", buffer, g_iniPath.c_str());
}

void UpdateLayeredStyle(bool enable)
{
    if (!g_hwnd)
    {
        return;
    }

    LONG_PTR exStyle = GetWindowLongPtr(g_hwnd, GWL_EXSTYLE);
    if (enable)
    {
        exStyle |= WS_EX_LAYERED;
    }
    else
    {
        exStyle &= ~WS_EX_LAYERED;
    }
    SetWindowLongPtr(g_hwnd, GWL_EXSTYLE, exStyle);
}

bool UpdateLayeredWindowFromWic(HWND hwnd, float drawWidth, float drawHeight)
{
    if (!g_wicSource || drawWidth <= 0.0f || drawHeight <= 0.0f)
    {
        return false;
    }

    UINT width = static_cast<UINT>(std::max(1.0f, static_cast<float>(std::lround(drawWidth))));
    UINT height = static_cast<UINT>(std::max(1.0f, static_cast<float>(std::lround(drawHeight))));

    IWICBitmapScaler* scaler = nullptr;
    HRESULT hr = g_wicFactory->CreateBitmapScaler(&scaler);
    if (FAILED(hr))
    {
        return false;
    }

    hr = scaler->Initialize(g_wicSource, width, height, WICBitmapInterpolationModeFant);
    if (FAILED(hr))
    {
        scaler->Release();
        return false;
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = static_cast<LONG>(width);
    bmi.bmiHeader.biHeight = -static_cast<LONG>(height);
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HDC screenDc = GetDC(nullptr);
    HBITMAP dib = CreateDIBSection(screenDc, &bmi, DIB_RGB_COLORS, &bits, nullptr, 0);
    HDC memDc = CreateCompatibleDC(screenDc);
    HGDIOBJ oldBmp = SelectObject(memDc, dib);

    WICRect rect{ 0, 0, static_cast<INT>(width), static_cast<INT>(height) };
    UINT stride = width * 4;
    UINT bufferSize = stride * height;
    hr = scaler->CopyPixels(&rect, stride, bufferSize, static_cast<BYTE*>(bits));

    POINT ptSrc{ 0, 0 };
    SIZE sizeWindow{ static_cast<LONG>(width), static_cast<LONG>(height) };
    RECT wndRect{};
    GetWindowRect(hwnd, &wndRect);
    POINT ptDst{ wndRect.left, wndRect.top };
    BLENDFUNCTION blend{};
    blend.BlendOp = AC_SRC_OVER;
    blend.SourceConstantAlpha = 255;
    blend.AlphaFormat = AC_SRC_ALPHA;

    bool updated = false;
    if (SUCCEEDED(hr))
    {
        updated = UpdateLayeredWindow(
            hwnd,
            screenDc,
            &ptDst,
            &sizeWindow,
            memDc,
            &ptSrc,
            0,
            &blend,
            ULW_ALPHA
        ) == TRUE;
    }

    SelectObject(memDc, oldBmp);
    DeleteDC(memDc);
    DeleteObject(dib);
    ReleaseDC(nullptr, screenDc);
    scaler->Release();

    return updated;
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

        if (g_imageHasAlpha)
        {
            UpdateLayeredWindowFromWic(hwnd, drawWidth, drawHeight);
            return;
        }

        if (g_renderTarget && drawWidth > 0.0f && drawHeight > 0.0f)
        {
            D2D1_SIZE_F rtSize = g_renderTarget->GetSize();
            float roundedWidth = static_cast<float>(std::lround(drawWidth));
            float roundedHeight = static_cast<float>(std::lround(drawHeight));
            UINT targetWidth = static_cast<UINT>(std::max(1.0f, roundedWidth));
            UINT targetHeight = static_cast<UINT>(std::max(1.0f, roundedHeight));
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
    SetProcessDPIAware();
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    const wchar_t CLASS_NAME[] = L"FloatVisionWindow";

    WNDCLASSEX wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_FLOATVISION));
    wc.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL));

    RegisterClassEx(&wc);

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
    if (g_hwnd)
    {
        wchar_t modulePath[MAX_PATH]{};
        if (GetModuleFileNameW(nullptr, modulePath, MAX_PATH) > 0)
        {
            std::filesystem::path iniPath(modulePath);
            iniPath.replace_extension(L".ini");
            g_iniPath = iniPath.wstring();
        }
    }

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

    LoadSettings();
    LoadWindowPlacement();
    ApplyAlwaysOnTop();
    if (g_hasSavedWindowPos)
    {
        SetWindowPos(
            hwnd,
            nullptr,
            g_windowPos.x,
            g_windowPos.y,
            0,
            0,
            SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE
        );
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
            UpdateZoomToFitScreen(hwnd);
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
    else
    {
        UpdateLayeredStyle(g_imageHasAlpha);
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
