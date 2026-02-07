#include <windows.h>
#include <d2d1.h>
#include <wincodec.h>
#include <dwrite.h>
#include <commdlg.h>
#include <algorithm>
#include <string>
#include <vector>

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

IWICImagingFactory* g_wicFactory = nullptr;
IDWriteFactory* g_dwriteFactory = nullptr;
IDWriteTextFormat* g_placeholderFormat = nullptr;

UINT g_imageWidth = 0;
UINT g_imageHeight = 0;

enum class SortMode
{
    NameAsc,
    NameDesc,
    ModifiedAsc,
    ModifiedDesc
};

struct ImageEntry
{
    std::wstring path;
    std::wstring name;
    FILETIME writeTime{};
};

struct ImageListState
{
    std::wstring currentPath;
    std::wstring directory;
    std::vector<ImageEntry> entries;
    size_t currentIndex = 0;
    SortMode sortMode = SortMode::NameAsc;
};

ImageListState g_imageList;

constexpr UINT ID_MENU_OPEN = 1001;
constexpr UINT ID_MENU_PREV = 1002;
constexpr UINT ID_MENU_NEXT = 1003;
constexpr UINT ID_MENU_SORT_NAME_ASC = 1101;
constexpr UINT ID_MENU_SORT_NAME_DESC = 1102;
constexpr UINT ID_MENU_SORT_MODIFIED_ASC = 1103;
constexpr UINT ID_MENU_SORT_MODIFIED_DESC = 1104;

// =====================
// 前方宣言
// =====================
void Render();
bool InitDirect2D(HWND hwnd);
bool InitWIC();
bool InitDirectWrite();
bool LoadImageFromFile(const wchar_t* path);
bool LoadImageAndUpdateList(HWND hwnd, const std::wstring& path);

static std::wstring ToLower(const std::wstring& text)
{
    std::wstring lowered = text;
    std::transform(lowered.begin(), lowered.end(), lowered.begin(), towlower);
    return lowered;
}

static bool HasImageExtension(const std::wstring& filename)
{
    const std::wstring lowered = ToLower(filename);
    const wchar_t* extensions[] = {
        L".bmp", L".png", L".jpg", L".jpeg", L".gif", L".tif", L".tiff", L".webp"
    };
    for (const auto& ext : extensions)
    {
        if (lowered.size() >= wcslen(ext) &&
            lowered.compare(lowered.size() - wcslen(ext), wcslen(ext), ext) == 0)
        {
            return true;
        }
    }
    return false;
}

static std::wstring GetDirectoryFromPath(const std::wstring& path)
{
    size_t pos = path.find_last_of(L"\\/");
    if (pos == std::wstring::npos)
    {
        return L"";
    }
    return path.substr(0, pos);
}

static void SortImageEntries()
{
    switch (g_imageList.sortMode)
    {
    case SortMode::NameAsc:
        std::sort(g_imageList.entries.begin(), g_imageList.entries.end(),
            [](const ImageEntry& a, const ImageEntry& b)
            {
                return _wcsicmp(a.name.c_str(), b.name.c_str()) < 0;
            });
        break;
    case SortMode::NameDesc:
        std::sort(g_imageList.entries.begin(), g_imageList.entries.end(),
            [](const ImageEntry& a, const ImageEntry& b)
            {
                return _wcsicmp(a.name.c_str(), b.name.c_str()) > 0;
            });
        break;
    case SortMode::ModifiedAsc:
        std::sort(g_imageList.entries.begin(), g_imageList.entries.end(),
            [](const ImageEntry& a, const ImageEntry& b)
            {
                return CompareFileTime(&a.writeTime, &b.writeTime) < 0;
            });
        break;
    case SortMode::ModifiedDesc:
        std::sort(g_imageList.entries.begin(), g_imageList.entries.end(),
            [](const ImageEntry& a, const ImageEntry& b)
            {
                return CompareFileTime(&a.writeTime, &b.writeTime) > 0;
            });
        break;
    }
}

static void RefreshImageListFromPath(const std::wstring& path)
{
    g_imageList.currentPath = path;
    g_imageList.directory = GetDirectoryFromPath(path);
    g_imageList.entries.clear();

    if (g_imageList.directory.empty())
    {
        g_imageList.currentIndex = 0;
        return;
    }

    std::wstring searchPattern = g_imageList.directory + L"\\*";
    WIN32_FIND_DATAW findData{};
    HANDLE findHandle = FindFirstFileW(searchPattern.c_str(), &findData);
    if (findHandle == INVALID_HANDLE_VALUE)
    {
        g_imageList.currentIndex = 0;
        return;
    }

    do
    {
        if (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        {
            continue;
        }
        const std::wstring filename = findData.cFileName;
        if (!HasImageExtension(filename))
        {
            continue;
        }
        ImageEntry entry;
        entry.name = filename;
        entry.path = g_imageList.directory + L"\\" + filename;
        entry.writeTime = findData.ftLastWriteTime;
        g_imageList.entries.push_back(entry);
    } while (FindNextFileW(findHandle, &findData));

    FindClose(findHandle);

    SortImageEntries();

    g_imageList.currentIndex = 0;
    for (size_t i = 0; i < g_imageList.entries.size(); ++i)
    {
        if (_wcsicmp(g_imageList.entries[i].path.c_str(), path.c_str()) == 0)
        {
            g_imageList.currentIndex = i;
            break;
        }
    }
}

static void MoveImageIndex(HWND hwnd, int delta)
{
    if (g_imageList.entries.empty())
    {
        return;
    }
    const int count = static_cast<int>(g_imageList.entries.size());
    int nextIndex = static_cast<int>(g_imageList.currentIndex) + delta;
    nextIndex = (nextIndex % count + count) % count;
    g_imageList.currentIndex = static_cast<size_t>(nextIndex);
    LoadImageAndUpdateList(hwnd, g_imageList.entries[g_imageList.currentIndex].path);
}

static void UpdateSortMode(HWND hwnd, SortMode mode)
{
    if (g_imageList.entries.empty())
    {
        g_imageList.sortMode = mode;
        return;
    }
    g_imageList.sortMode = mode;
    const std::wstring currentPath = g_imageList.currentPath;
    SortImageEntries();
    for (size_t i = 0; i < g_imageList.entries.size(); ++i)
    {
        if (_wcsicmp(g_imageList.entries[i].path.c_str(), currentPath.c_str()) == 0)
        {
            g_imageList.currentIndex = i;
            break;
        }
    }
    InvalidateRect(hwnd, nullptr, TRUE);
}

static void OpenImageFileDialog(HWND hwnd)
{
    wchar_t fileName[MAX_PATH] = {};
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFilter = L"Image Files\0*.bmp;*.png;*.jpg;*.jpeg;*.gif;*.tif;*.tiff;*.webp\0All Files\0*.*\0";
    ofn.lpstrFile = fileName;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    if (!g_imageList.directory.empty())
    {
        ofn.lpstrInitialDir = g_imageList.directory.c_str();
    }

    if (GetOpenFileNameW(&ofn))
    {
        LoadImageAndUpdateList(hwnd, fileName);
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
        Render();
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
        return 0;
    }

    case WM_CONTEXTMENU:
    {
        POINT pt{
            GET_X_LPARAM(lParam),
            GET_Y_LPARAM(lParam)
        };
        if (pt.x == -1 && pt.y == -1)
        {
            GetCursorPos(&pt);
        }

        HMENU menu = CreatePopupMenu();
        if (!menu)
        {
            return 0;
        }

        AppendMenuW(menu, MF_STRING, ID_MENU_OPEN, L"ファイルを開く");
        AppendMenuW(menu, MF_STRING, ID_MENU_PREV, L"前の画像");
        AppendMenuW(menu, MF_STRING, ID_MENU_NEXT, L"次の画像");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, ID_MENU_SORT_NAME_ASC, L"名前 昇順");
        AppendMenuW(menu, MF_STRING, ID_MENU_SORT_NAME_DESC, L"名前 降順");
        AppendMenuW(menu, MF_STRING, ID_MENU_SORT_MODIFIED_ASC, L"更新日時 昇順");
        AppendMenuW(menu, MF_STRING, ID_MENU_SORT_MODIFIED_DESC, L"更新日時 降順");

        const UINT command = TrackPopupMenu(
            menu,
            TPM_RETURNCMD | TPM_RIGHTBUTTON,
            pt.x,
            pt.y,
            0,
            hwnd,
            nullptr
        );
        DestroyMenu(menu);

        if (command != 0)
        {
            SendMessage(hwnd, WM_COMMAND, command, 0);
        }
        return 0;
    }

    case WM_COMMAND:
    {
        switch (LOWORD(wParam))
        {
        case ID_MENU_OPEN:
            OpenImageFileDialog(hwnd);
            break;
        case ID_MENU_PREV:
            MoveImageIndex(hwnd, -1);
            break;
        case ID_MENU_NEXT:
            MoveImageIndex(hwnd, 1);
            break;
        case ID_MENU_SORT_NAME_ASC:
            UpdateSortMode(hwnd, SortMode::NameAsc);
            break;
        case ID_MENU_SORT_NAME_DESC:
            UpdateSortMode(hwnd, SortMode::NameDesc);
            break;
        case ID_MENU_SORT_MODIFIED_ASC:
            UpdateSortMode(hwnd, SortMode::ModifiedAsc);
            break;
        case ID_MENU_SORT_MODIFIED_DESC:
            UpdateSortMode(hwnd, SortMode::ModifiedDesc);
            break;
        default:
            break;
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

bool LoadImageAndUpdateList(HWND hwnd, const std::wstring& path)
{
    if (!LoadImageFromFile(path.c_str()))
    {
        return false;
    }
    RefreshImageListFromPath(path);
    InvalidateRect(hwnd, nullptr, TRUE);
    return true;
}

// =====================
// 描画
// =====================
void Render()
{
    if (!g_renderTarget)
        return;

    g_renderTarget->BeginDraw();

    if (g_bitmap)
    {
        g_renderTarget->Clear(D2D1::ColorF(D2D1::ColorF::Black));
        D2D1_SIZE_F rtSize = g_renderTarget->GetSize();

        float x = (rtSize.width - g_imageWidth) * 0.5f;
        float y = (rtSize.height - g_imageHeight) * 0.5f;

        D2D1_RECT_F dest = D2D1::RectF(
            x,
            y,
            x + g_imageWidth,
            y + g_imageHeight
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

    g_renderTarget->EndDraw();
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

    if (!hwnd)
    {
        MessageBox(nullptr, L"CreateWindowEx failed", L"Error", MB_OK);
        return 0;
    }

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

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    if (!InitDirect2D(hwnd))
    {
        MessageBox(hwnd, L"Direct2D init failed", L"Error", MB_OK);
        return 0;
    }

    int argc = 0;
    wchar_t** argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv && argc > 1)
    {
        LoadImageAndUpdateList(hwnd, argv[1]);
    }
    if (argv)
    {
        LocalFree(argv);
    }
    InvalidateRect(hwnd, nullptr, TRUE);

    MSG msg{};
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}
