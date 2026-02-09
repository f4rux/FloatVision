#define NOMINMAX
#include <windows.h>
#include <windowsx.h>
#include <d2d1.h>
#include <wincodec.h>
#include <dwrite.h>
#include <shellapi.h>
#include <commdlg.h>
#include <commctrl.h>
#include <uxtheme.h>
#include <dwmapi.h>
#include <filesystem>
#include <vector>
#include <array>
#include <algorithm>
#include <cmath>
#include <cstdlib>
#include <cwctype>
#include <fstream>
#include <string>
#include <wrl.h>
#include <WebView2.h>
#include "resource.h"
#include "md4c.h"
#include "md4c-html.h"
#include "entity.h"
#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "windowscodecs.lib")
#pragma comment(lib, "dwrite.lib")
#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "UxTheme.lib")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "Advapi32.lib")
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#ifndef FLOATVISION_GLOBALS_DEFINED
#define FLOATVISION_GLOBALS_DEFINED
// =====================
// グローバル変数
// =====================
ID2D1Factory* g_d2dFactory = nullptr;
ID2D1HwndRenderTarget* g_renderTarget = nullptr;
ID2D1Bitmap* g_bitmap = nullptr;
ID2D1SolidColorBrush* g_placeholderBrush = nullptr;
ID2D1SolidColorBrush* g_checkerBrushA = nullptr;
ID2D1SolidColorBrush* g_checkerBrushB = nullptr;
ID2D1SolidColorBrush* g_customColorBrush = nullptr;
ID2D1SolidColorBrush* g_textBrush = nullptr;
HWND g_hwnd = nullptr;
IWICImagingFactory* g_wicFactory = nullptr;
IDWriteFactory* g_dwriteFactory = nullptr;
IDWriteTextFormat* g_placeholderFormat = nullptr;
IDWriteTextFormat* g_textFormat = nullptr;
IWICBitmapSource* g_wicSourceStraight = nullptr;
IWICBitmapSource* g_wicSourcePremultiplied = nullptr;
Microsoft::WRL::ComPtr<ICoreWebView2Controller> g_webviewController;
Microsoft::WRL::ComPtr<ICoreWebView2Controller2> g_webviewController2;
Microsoft::WRL::ComPtr<ICoreWebView2> g_webview;
HMODULE g_webviewLoader = nullptr;
HWND g_webviewWindow = nullptr;

UINT g_imageWidth = 0;
UINT g_imageHeight = 0;
bool g_imageHasAlpha = false;

float g_zoom = 1.0f;
bool g_fitToWindow = true;
bool g_isEdgeDragging = false;
POINT g_dragStartPoint{};
float g_dragStartZoom = 1.0f;
float g_dragStartScale = 1.0f;
float g_dragStartWidth = 0.0f;
float g_dragStartHeight = 0.0f;

bool g_hasText = false;
std::wstring g_textContent;
std::wstring g_textFontName = L"Consolas";
float g_textFontSize = 18.0f;
COLORREF g_textColor = RGB(240, 240, 240);
COLORREF g_textBackground = RGB(20, 20, 20);
bool g_textWrap = true;
float g_textScroll = 0.0f;
bool g_hasHtml = false;
std::wstring g_pendingHtmlContent;
bool g_webviewInputTimerActive = false;
enum class HtmlInputKey
{
    Shift = 0,
    Ctrl = 1,
    Alt = 2
};
HtmlInputKey g_htmlInputKey = HtmlInputKey::Alt;
WORD g_keyNextFile = VK_RIGHT;
WORD g_keyPrevFile = VK_LEFT;
WORD g_keyZoomIn = VK_UP;
WORD g_keyZoomOut = VK_DOWN;
WORD g_keyOpenFile = 'O';
WORD g_keyExit = VK_ESCAPE;
WORD g_keyAlwaysOnTop = 'P';

enum class TransparencyMode
{
    Transparent = 0,
    Checkerboard = 1,
    SolidColor = 2
};

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
bool g_sortImageOnly = true;
std::filesystem::path g_currentImagePath;
const float g_zoomMin = 0.05f;
const float g_zoomMax = 20.0f;
const float g_edgeDragMargin = 12.0f;
bool g_alwaysOnTop = false;
std::wstring g_iniPath;
POINT g_windowPos{ CW_USEDEFAULT, CW_USEDEFAULT };
bool g_hasSavedWindowPos = false;
TransparencyMode g_transparencyMode = TransparencyMode::Transparent;
COLORREF g_customColor = RGB(32, 32, 32);

constexpr int kMenuOpen = 1001;
constexpr int kMenuNext = 1002;
constexpr int kMenuPrev = 1003;
constexpr int kMenuZoomIn = 1005;
constexpr int kMenuZoomOut = 1006;
constexpr int kMenuOriginalSize = 1007;
constexpr int kMenuAlwaysOnTop = 1008;
constexpr int kMenuExit = 1009;
constexpr int kMenuSettings = 1010;
constexpr int kMenuSortNameAsc = 1101;
constexpr int kMenuSortNameDesc = 1102;
constexpr int kMenuSortTimeAsc = 1103;
constexpr int kMenuSortTimeDesc = 1104;
constexpr int kMenuSortImageOnly = 1105;
constexpr UINT_PTR kWebViewInputTimerId = 2001;
constexpr UINT kWebViewInputTimerIntervalMs = 50;

// =====================
// 前方宣言
// =====================
void Render(HWND hwnd);
bool InitDirect2D(HWND hwnd);
bool InitWIC();
bool InitDirectWrite();
bool LoadImageFromFile(const wchar_t* path);
bool LoadTextFromFile(const wchar_t* path);
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
bool ImageHasTransparency(IWICBitmapSource* source);
void ApplyTransparencyMode();
void UpdateCustomColorBrush();
void ShowSettingsDialog(HWND hwnd);
void UpdateTextFormat();
void UpdateTextBrush();
void ResizeWindowByFactor(HWND hwnd, float factor);
void ScrollTextBy(float delta);
bool LoadHtmlFromFile(const wchar_t* path);
bool LoadMarkdownFromFile(const wchar_t* path);
std::wstring InjectHtmlBaseStyles(const std::wstring& html);
bool ReadFileBytes(const wchar_t* path, std::string& bytes);
bool Utf8ToWide(const std::string& bytes, std::wstring& text);
bool ApplyHtmlContent(std::wstring html);
bool RenderMarkdownToHtml(const std::string& markdown, std::string& html);
void UpdateWebViewInputTimer();
WORD GetHtmlInputVirtualKey();
void UpdateWebViewInputState();
void UpdateWebViewWindowHandle();
bool EnsureWebView2(HWND hwnd);
void UpdateWebViewBounds();
void HideWebView();
void CloseWebView();
void RefreshMenuTheme();
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

bool IsTextFile(const std::filesystem::path& path)
{
    if (!path.has_extension())
    {
        return false;
    }
    std::wstring ext = path.extension().wstring();
    std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
    return ext == L".txt";
}

bool IsHtmlFile(const std::filesystem::path& path)
{
    if (!path.has_extension())
    {
        return false;
    }
    std::wstring ext = path.extension().wstring();
    std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
    return ext == L".html" || ext == L".htm";
}

bool IsMarkdownFile(const std::filesystem::path& path)
{
    if (!path.has_extension())
    {
        return false;
    }
    std::wstring ext = path.extension().wstring();
    std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
    return ext == L".md" || ext == L".markdown";
}

bool IsSupportedFile(const std::filesystem::path& path)
{
    return IsImageFile(path) || IsTextFile(path) || IsHtmlFile(path) || IsMarkdownFile(path);
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
        if (g_webviewController)
        {
            UpdateWebViewBounds();
        }
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
                if (IsMarkdownFile(path))
                {
                    if (LoadMarkdownFromFile(path.c_str()))
                    {
                        RefreshImageList(path);
                        InvalidateRect(hwnd, nullptr, TRUE);
                    }
                }
                else if (IsHtmlFile(path))
                {
                    if (LoadHtmlFromFile(path.c_str()))
                    {
                        RefreshImageList(path);
                        InvalidateRect(hwnd, nullptr, TRUE);
                    }
                }
                else if (IsTextFile(path))
                {
                    if (LoadTextFromFile(path.c_str()))
                    {
                        RefreshImageList(path);
                        InvalidateRect(hwnd, nullptr, TRUE);
                    }
                }
                else if (LoadImageFromFile(path.c_str()))
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
        AppendMenu(menu, MF_STRING, kMenuSettings, L"Settings...");
        AppendMenu(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenu(menu, MF_STRING, kMenuSortNameAsc, L"Sort: Name (A-Z)");
        AppendMenu(menu, MF_STRING, kMenuSortNameDesc, L"Sort: Name (Z-A)");
        AppendMenu(menu, MF_STRING, kMenuSortTimeAsc, L"Sort: Modified (Old-New)");
        AppendMenu(menu, MF_STRING, kMenuSortTimeDesc, L"Sort: Modified (New-Old)");
        AppendMenu(menu, MF_STRING, kMenuSortImageOnly, L"Sort: Image only");
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
        CheckMenuItem(menu, kMenuSortImageOnly, MF_BYCOMMAND | (g_sortImageOnly ? MF_CHECKED : MF_UNCHECKED));
        CheckMenuItem(menu, kMenuAlwaysOnTop, MF_BYCOMMAND | (g_alwaysOnTop ? MF_CHECKED : MF_UNCHECKED));

        POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        RefreshMenuTheme();
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
                if (g_bitmap && !g_hasHtml)
                {
                    UpdateZoomToFitScreen(hwnd);
                }
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
        case kMenuSettings:
            ShowSettingsDialog(hwnd);
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
        case kMenuSortImageOnly:
            g_sortImageOnly = !g_sortImageOnly;
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
        if (g_hasHtml)
        {
            WORD inputKey = GetHtmlInputVirtualKey();
            bool keyDown = (GetKeyState(inputKey) & 0x8000) != 0;
            if (keyDown && g_webviewWindow)
            {
                WPARAM adjustedWParam = wParam;
                if (inputKey == VK_SHIFT)
                {
                    WORD keyState = GET_KEYSTATE_WPARAM(wParam);
                    adjustedWParam = MAKEWPARAM(keyState & ~MK_SHIFT, GET_WHEEL_DELTA_WPARAM(wParam));
                }
                SendMessageW(g_webviewWindow, WM_MOUSEWHEEL, adjustedWParam, lParam);
                return 0;
            }
            return DefWindowProc(hwnd, msg, wParam, lParam);
        }
        if (!g_bitmap && !g_hasText)
        {
            return 0;
        }
        int delta = GET_WHEEL_DELTA_WPARAM(wParam);
        float steps = static_cast<float>(delta) / WHEEL_DELTA;
        if (g_hasText)
        {
            ScrollTextBy(-steps * 40.0f);
            InvalidateRect(hwnd, nullptr, TRUE);
        }
        else
        {
            float factor = std::pow(1.1f, steps);
            POINT pt{};
            GetCursorPos(&pt);
            AdjustZoom(factor, pt);
            InvalidateRect(hwnd, nullptr, TRUE);
        }
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
        if (nearEdge && (g_bitmap || g_hasText || g_hasHtml))
        {
            g_fitToWindow = false;
            g_isEdgeDragging = true;
            g_dragStartPoint = pt;
            g_dragStartZoom = g_zoom;
            g_dragStartScale = std::max(1.0f, (std::min)(static_cast<float>(rc.right - rc.left), static_cast<float>(rc.bottom - rc.top)));
            g_dragStartWidth = static_cast<float>(rc.right - rc.left);
            g_dragStartHeight = static_cast<float>(rc.bottom - rc.top);
            if (g_bitmap)
            {
                UpdateWindowToZoomedImage();
            }
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
            float deltaX = static_cast<float>(pt.x - g_dragStartPoint.x);
            float deltaY = static_cast<float>(pt.y - g_dragStartPoint.y);
            if (g_hasText || g_hasHtml)
            {
                float nextWidth = std::max(200.0f, g_dragStartWidth + deltaX);
                float nextHeight = std::max(200.0f, g_dragStartHeight + deltaY);
                SetWindowPos(hwnd, nullptr, 0, 0, static_cast<int>(nextWidth), static_cast<int>(nextHeight),
                    SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            else
            {
                float nextScale = std::max(1.0f, g_dragStartScale + deltaY);
                float zoom = (g_dragStartScale > 0.0f) ? (g_dragStartZoom * (nextScale / g_dragStartScale)) : g_dragStartZoom;
                g_zoom = std::max(g_zoomMin, (std::min)(zoom, g_zoomMax));
                UpdateWindowToZoomedImage();
                InvalidateRect(hwnd, nullptr, TRUE);
            }
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
        WORD key = static_cast<WORD>(wParam);
        if (key == g_keyExit)
        {
            DestroyWindow(hwnd);
            return 0;
        }
        if (key == g_keyAlwaysOnTop)
        {
            g_alwaysOnTop = !g_alwaysOnTop;
            ApplyAlwaysOnTop();
            SaveSettings();
            return 0;
        }
        if (key == g_keyOpenFile)
        {
            if (ShowOpenImageDialog(hwnd))
            {
                if (g_bitmap && !g_hasHtml)
                {
                    UpdateZoomToFitScreen(hwnd);
                }
                InvalidateRect(hwnd, nullptr, TRUE);
            }
            return 0;
        }
        if (g_hasHtml)
        {
            WORD inputKey = GetHtmlInputVirtualKey();
            if (wParam == inputKey)
            {
                UpdateWebViewInputState();
            }
            return DefWindowProc(hwnd, msg, wParam, lParam);
        }
        if ((key == g_keyZoomIn || key == g_keyZoomOut) && g_hasText)
        {
            float delta = (key == g_keyZoomIn) ? -40.0f : 40.0f;
            ScrollTextBy(delta);
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        }
        if (key == g_keyNextFile && !g_imageList.empty())
        {
            NavigateImage(1);
            return 0;
        }
        if (key == g_keyPrevFile && !g_imageList.empty())
        {
            NavigateImage(-1);
            return 0;
        }
        if ((key == g_keyZoomIn || key == g_keyZoomOut) && g_bitmap)
        {
            float factor = (key == g_keyZoomIn) ? 1.1f : (1.0f / 1.1f);
            POINT pt{};
            GetCursorPos(&pt);
            AdjustZoom(factor, pt);
            InvalidateRect(hwnd, nullptr, TRUE);
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

    case WM_KEYUP:
    {
        if (g_hasHtml)
        {
            WORD inputKey = GetHtmlInputVirtualKey();
            if (wParam == inputKey)
            {
                UpdateWebViewInputState();
            }
            return DefWindowProc(hwnd, msg, wParam, lParam);
        }
        return 0;
    }

    case WM_TIMER:
    {
        if (wParam == kWebViewInputTimerId)
        {
            if (g_hasHtml)
            {
                UpdateWebViewInputState();
            }
            else
            {
                UpdateWebViewInputTimer();
            }
            return 0;
        }
        break;
    }

    case WM_DESTROY:
    {
        CloseWebView();
        SaveWindowPlacement();
        SaveSettings();
        PostQuitMessage(0);
        return 0;
    }

    default:
        break;
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
        D2D1::ColorF(0.992f, 0.992f, 0.992f),
        &g_placeholderBrush)))
    {
        return false;
    }

    UpdateCustomColorBrush();
    UpdateTextBrush();

    if (FAILED(g_renderTarget->CreateSolidColorBrush(
        D2D1::ColorF(0.85f, 0.85f, 0.85f),
        &g_checkerBrushA)))
    {
        return false;
    }
    if (FAILED(g_renderTarget->CreateSolidColorBrush(
        D2D1::ColorF(0.7f, 0.7f, 0.7f),
        &g_checkerBrushB)))
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
    if (g_checkerBrushA)
    {
        g_checkerBrushA->Release();
        g_checkerBrushA = nullptr;
    }
    if (g_checkerBrushB)
    {
        g_checkerBrushB->Release();
        g_checkerBrushB = nullptr;
    }
    if (g_customColorBrush)
    {
        g_customColorBrush->Release();
        g_customColorBrush = nullptr;
    }
    if (g_textBrush)
    {
        g_textBrush->Release();
        g_textBrush = nullptr;
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
    if (g_wicSourceStraight)
    {
        g_wicSourceStraight->Release();
        g_wicSourceStraight = nullptr;
    }
    if (g_wicSourcePremultiplied)
    {
        g_wicSourcePremultiplied->Release();
        g_wicSourcePremultiplied = nullptr;
    }
    g_imageWidth = 0;
    g_imageHeight = 0;
    g_imageHasAlpha = false;
}

void CleanupResources()
{
    DiscardRenderTarget();
    CloseWebView();

    if (g_placeholderFormat)
    {
        g_placeholderFormat->Release();
        g_placeholderFormat = nullptr;
    }
    if (g_textFormat)
    {
        g_textFormat->Release();
        g_textFormat = nullptr;
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

    UpdateTextFormat();

    return true;
}

// =====================
// 画像ロード
// =====================
bool LoadImageFromFile(const wchar_t* path)
{
    IWICBitmapDecoder* decoder = nullptr;
    IWICBitmapFrameDecode* frame = nullptr;
    IWICFormatConverter* converterStraight = nullptr;
    IWICFormatConverter* converterPremultiplied = nullptr;
    WICPixelFormatGUID pixelFormat = GUID_WICPixelFormatDontCare;
    D2D1_BITMAP_PROPERTIES bitmapProperties{};

    if (g_bitmap)
    {
        g_bitmap->Release();
        g_bitmap = nullptr;
    }
    if (g_wicSourceStraight)
    {
        g_wicSourceStraight->Release();
        g_wicSourceStraight = nullptr;
    }
    if (g_wicSourcePremultiplied)
    {
        g_wicSourcePremultiplied->Release();
        g_wicSourcePremultiplied = nullptr;
    }
    g_imageWidth = 0;
    g_imageHeight = 0;
    g_imageHasAlpha = false;
    g_hasText = false;
    g_textContent.clear();
    g_hasHtml = false;
    g_pendingHtmlContent.clear();
    HideWebView();

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

    hr = g_wicFactory->CreateFormatConverter(&converterStraight);
    if (FAILED(hr)) goto cleanup;

    hr = converterStraight->Initialize(
        frame,
        GUID_WICPixelFormat32bppBGRA,
        WICBitmapDitherTypeNone,
        nullptr,
        0.0,
        WICBitmapPaletteTypeCustom
    );
    if (FAILED(hr)) goto cleanup;

    if (!g_imageHasAlpha)
    {
        g_imageHasAlpha = ImageHasTransparency(converterStraight);
    }

    hr = g_wicFactory->CreateFormatConverter(&converterPremultiplied);
    if (FAILED(hr)) goto cleanup;

    hr = converterPremultiplied->Initialize(
        converterStraight,
        GUID_WICPixelFormat32bppPBGRA,
        WICBitmapDitherTypeNone,
        nullptr,
        0.0,
        WICBitmapPaletteTypeCustom
    );
    if (FAILED(hr)) goto cleanup;

    bitmapProperties = D2D1::BitmapProperties(
        D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED)
    );

    hr = g_renderTarget->CreateBitmapFromWicBitmap(
        converterPremultiplied,
        &bitmapProperties,
        &g_bitmap
    );
    if (FAILED(hr))
    {
        bitmapProperties = D2D1::BitmapProperties(
            D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED)
        );
        hr = g_renderTarget->CreateBitmapFromWicBitmap(
            converterPremultiplied,
            &bitmapProperties,
            &g_bitmap
        );
    }
    if (SUCCEEDED(hr))
    {
        g_wicSourceStraight = converterStraight;
        g_wicSourceStraight->AddRef();
        g_wicSourcePremultiplied = converterPremultiplied;
        g_wicSourcePremultiplied->AddRef();
    }

cleanup:
    if (decoder) decoder->Release();
    if (frame) frame->Release();
    if (converterStraight) converterStraight->Release();
    if (converterPremultiplied) converterPremultiplied->Release();

    ApplyTransparencyMode();
    return SUCCEEDED(hr);
}

bool ReadFileBytes(const wchar_t* path, std::string& bytes)
{
    std::ifstream file(path, std::ios::binary);
    if (!file)
    {
        return false;
    }

    bytes.assign((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
    if (bytes.size() >= 3 && static_cast<unsigned char>(bytes[0]) == 0xEF
        && static_cast<unsigned char>(bytes[1]) == 0xBB
        && static_cast<unsigned char>(bytes[2]) == 0xBF)
    {
        bytes.erase(0, 3);
    }
    return true;
}

bool Utf8ToWide(const std::string& bytes, std::wstring& text)
{
    if (bytes.empty())
    {
        text.clear();
        return true;
    }

    int needed = MultiByteToWideChar(CP_UTF8, 0, bytes.data(), static_cast<int>(bytes.size()), nullptr, 0);
    if (needed <= 0)
    {
        return false;
    }

    text.assign(needed, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, bytes.data(), static_cast<int>(bytes.size()), text.data(), needed);
    return true;
}

bool LoadTextFromFile(const wchar_t* path)
{
    std::string bytes;
    if (!ReadFileBytes(path, bytes))
    {
        return false;
    }

    std::wstring text;
    if (!Utf8ToWide(bytes, text))
    {
        return false;
    }

    if (g_bitmap)
    {
        g_bitmap->Release();
        g_bitmap = nullptr;
    }
    if (g_wicSourceStraight)
    {
        g_wicSourceStraight->Release();
        g_wicSourceStraight = nullptr;
    }
    if (g_wicSourcePremultiplied)
    {
        g_wicSourcePremultiplied->Release();
        g_wicSourcePremultiplied = nullptr;
    }
    g_imageWidth = 0;
    g_imageHeight = 0;
    g_imageHasAlpha = false;
    g_hasHtml = false;
    g_pendingHtmlContent.clear();
    g_hasText = true;
    g_textContent = std::move(text);
    g_textScroll = 0.0f;
    HideWebView();
    ApplyTransparencyMode();
    return true;
}

std::wstring InjectHtmlBaseStyles(const std::wstring& html)
{
    const std::wstring style = L"<style>html, body { background: #ffffff !important; margin: 0; width: 100%; height: 100%; overflow: auto; }</style>";
    std::wstring lowered = html;
    std::transform(lowered.begin(), lowered.end(), lowered.begin(), ::towlower);
    size_t headPos = lowered.find(L"<head");
    if (headPos != std::wstring::npos)
    {
        size_t insertPos = lowered.find(L'>', headPos);
        if (insertPos != std::wstring::npos)
        {
            std::wstring result = html;
            result.insert(insertPos + 1, style);
            return result;
        }
    }
    return style + html;
}

bool RenderMarkdownToHtml(const std::string& markdown, std::string& html)
{
    std::string body;
    auto appendChunk = [](const MD_CHAR* text, MD_SIZE size, void* userdata)
    {
        auto* output = static_cast<std::string*>(userdata);
        output->append(text, size);
    };

    if (md_html(markdown.data(), static_cast<MD_SIZE>(markdown.size()), appendChunk, &body, MD_DIALECT_GITHUB, 0) != 0)
    {
        return false;
    }

    const char* kMarkdownStyle = R"(
        :root {
            color-scheme: light dark;
        }
        body {
            margin: 0;
            padding: 24px;
            font-family: "Segoe UI", "Meiryo", sans-serif;
            background: #ffffff;
            color: #1f2328;
        }
        .markdown-body {
            max-width: 960px;
            margin: 0 auto;
        }
        pre, code {
            font-family: "Consolas", "Courier New", monospace;
        }
        pre {
            padding: 12px;
            background: #f6f8fa;
            border-radius: 6px;
            overflow: auto;
        }
        code {
            background: #f6f8fa;
            border-radius: 4px;
            padding: 0 4px;
        }
        table {
            border-collapse: collapse;
        }
        th, td {
            border: 1px solid #d0d7de;
            padding: 6px 10px;
        }
        blockquote {
            border-left: 4px solid #d0d7de;
            margin: 0;
            padding-left: 16px;
            color: #57606a;
        }
    )";

    html = "<!DOCTYPE html><html><head><meta charset=\"utf-8\"><style>";
    html += kMarkdownStyle;
    html += "</style></head><body><article class=\"markdown-body\">";
    html += body;
    html += "</article></body></html>";
    return true;
}

bool ApplyHtmlContent(std::wstring html)
{
    if (g_bitmap)
    {
        g_bitmap->Release();
        g_bitmap = nullptr;
    }
    if (g_wicSourceStraight)
    {
        g_wicSourceStraight->Release();
        g_wicSourceStraight = nullptr;
    }
    if (g_wicSourcePremultiplied)
    {
        g_wicSourcePremultiplied->Release();
        g_wicSourcePremultiplied = nullptr;
    }
    g_imageWidth = 0;
    g_imageHeight = 0;
    g_imageHasAlpha = false;
    g_hasText = false;
    g_textContent.clear();
    g_textScroll = 0.0f;
    g_fitToWindow = false;
    g_zoom = 1.0f;
    g_hasHtml = true;
    g_pendingHtmlContent = std::move(html);
    ApplyTransparencyMode();

    if (!EnsureWebView2(g_hwnd))
    {
        g_hasHtml = false;
        g_pendingHtmlContent.clear();
        return false;
    }

    return true;
}

bool LoadHtmlFromFile(const wchar_t* path)
{
    std::string bytes;
    if (!ReadFileBytes(path, bytes))
    {
        return false;
    }

    std::wstring html;
    if (!Utf8ToWide(bytes, html))
    {
        return false;
    }

    html = InjectHtmlBaseStyles(html);
    return ApplyHtmlContent(std::move(html));
}

bool LoadMarkdownFromFile(const wchar_t* path)
{
    std::string bytes;
    if (!ReadFileBytes(path, bytes))
    {
        return false;
    }

    std::string htmlUtf8;
    if (!RenderMarkdownToHtml(bytes, htmlUtf8))
    {
        return false;
    }

    std::wstring html;
    if (!Utf8ToWide(htmlUtf8, html))
    {
        return false;
    }

    return ApplyHtmlContent(std::move(html));
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

bool ImageHasTransparency(IWICBitmapSource* source)
{
    if (!source)
    {
        return false;
    }

    UINT width = 0;
    UINT height = 0;
    HRESULT hr = source->GetSize(&width, &height);
    if (FAILED(hr) || width == 0 || height == 0)
    {
        return false;
    }

    const UINT stride = width * 4;
    std::vector<BYTE> row(stride);
    WICRect rect{ 0, 0, static_cast<INT>(width), 1 };

    for (UINT y = 0; y < height; ++y)
    {
        rect.Y = static_cast<INT>(y);
        hr = source->CopyPixels(&rect, stride, stride, row.data());
        if (FAILED(hr))
        {
            return false;
        }
        for (UINT x = 0; x < width; ++x)
        {
            BYTE alpha = row[x * 4 + 3];
            if (alpha < 255)
            {
                return true;
            }
        }
    }

    return false;
}

void ApplyTransparencyMode()
{
    if (g_imageHasAlpha && g_transparencyMode == TransparencyMode::Transparent)
    {
        UpdateLayeredStyle(true);
    }
    else
    {
        UpdateLayeredStyle(false);
    }
    UpdateCustomColorBrush();
}

void HideWebView()
{
    if (g_webviewController)
    {
        g_webviewController->put_IsVisible(FALSE);
    }
    if (g_webviewInputTimerActive && g_hwnd)
    {
        KillTimer(g_hwnd, kWebViewInputTimerId);
        g_webviewInputTimerActive = false;
    }
}

void UpdateWebViewBounds()
{
    if (!g_webviewController || !g_hwnd)
    {
        return;
    }
    RECT bounds{};
    GetClientRect(g_hwnd, &bounds);
    g_webviewController->put_Bounds(bounds);
}

void UpdateWebViewWindowHandle()
{
    if (!g_hwnd)
    {
        g_webviewWindow = nullptr;
        return;
    }

    HWND found = nullptr;
    EnumChildWindows(
        g_hwnd,
        [](HWND child, LPARAM lParam) -> BOOL
        {
            wchar_t className[64] = {};
            GetClassNameW(child, className, static_cast<int>(std::size(className)));
            if (wcsncmp(className, L"Chrome_WidgetWin", 16) == 0)
            {
                *reinterpret_cast<HWND*>(lParam) = child;
                return FALSE;
            }
            return TRUE;
        },
        reinterpret_cast<LPARAM>(&found));

    if (!found)
    {
        found = GetWindow(g_hwnd, GW_CHILD);
    }
    g_webviewWindow = found;
}

void UpdateWebViewInputTimer()
{
    if (!g_hwnd)
    {
        return;
    }
    if (g_hasHtml)
    {
        if (!g_webviewInputTimerActive)
        {
            SetTimer(g_hwnd, kWebViewInputTimerId, kWebViewInputTimerIntervalMs, nullptr);
            g_webviewInputTimerActive = true;
        }
    }
    else if (g_webviewInputTimerActive)
    {
        KillTimer(g_hwnd, kWebViewInputTimerId);
        g_webviewInputTimerActive = false;
    }
}

WORD GetHtmlInputVirtualKey()
{
    switch (g_htmlInputKey)
    {
    case HtmlInputKey::Ctrl:
        return VK_CONTROL;
    case HtmlInputKey::Alt:
        return VK_MENU;
    case HtmlInputKey::Shift:
    default:
        return VK_SHIFT;
    }
}

void UpdateWebViewInputState()
{
    if (!g_webviewWindow)
    {
        return;
    }

    LONG_PTR exStyle = GetWindowLongPtrW(g_webviewWindow, GWL_EXSTYLE);
    WORD inputKey = GetHtmlInputVirtualKey();
    bool keyDown = (GetKeyState(inputKey) & 0x8000) != 0;
    if (g_hasHtml && !keyDown)
    {
        exStyle |= WS_EX_TRANSPARENT;
        EnableWindow(g_webviewWindow, FALSE);
    }
    else
    {
        exStyle &= ~static_cast<LONG_PTR>(WS_EX_TRANSPARENT);
        EnableWindow(g_webviewWindow, TRUE);
    }
    SetWindowLongPtrW(g_webviewWindow, GWL_EXSTYLE, exStyle);
    SetWindowPos(
        g_webviewWindow,
        nullptr,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
}

bool EnsureWebView2(HWND hwnd)
{
    if (!hwnd)
    {
        return false;
    }

    if (g_webviewController && g_webview)
    {
        g_webviewController->put_IsVisible(TRUE);
        UpdateWebViewWindowHandle();
        UpdateWebViewInputState();
        UpdateWebViewInputTimer();
        UpdateWebViewBounds();
        if (!g_pendingHtmlContent.empty())
        {
            g_webview->NavigateToString(g_pendingHtmlContent.c_str());
            g_pendingHtmlContent.clear();
        }
        return true;
    }

    if (!g_webviewLoader)
    {
        g_webviewLoader = LoadLibraryW(L"WebView2Loader.dll");
        if (!g_webviewLoader)
        {
            MessageBoxW(hwnd, L"WebView2Loader.dll not found.", L"WebView2", MB_OK | MB_ICONERROR);
            return false;
        }
    }

    using CreateWebView2EnvironmentWithOptionsFn = HRESULT(WINAPI*)(
        PCWSTR, PCWSTR, ICoreWebView2EnvironmentOptions*,
        ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*);
    auto createEnv = reinterpret_cast<CreateWebView2EnvironmentWithOptionsFn>(
        GetProcAddress(g_webviewLoader, "CreateCoreWebView2EnvironmentWithOptions"));
    if (!createEnv)
    {
        MessageBoxW(hwnd, L"CreateCoreWebView2EnvironmentWithOptions not available.", L"WebView2", MB_OK | MB_ICONERROR);
        return false;
    }

    HRESULT hr = createEnv(
        nullptr,
        nullptr,
        nullptr,
        Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
            [hwnd](HRESULT result, ICoreWebView2Environment* env) -> HRESULT
            {
                if (FAILED(result) || !env)
                {
                    return result;
                }
                return env->CreateCoreWebView2Controller(
                    hwnd,
                    Microsoft::WRL::Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
                        [](HRESULT result, ICoreWebView2Controller* controller) -> HRESULT
                        {
                            if (FAILED(result) || !controller)
                            {
                                return result;
                            }
                            g_webviewController = controller;
                            g_webviewController->get_CoreWebView2(&g_webview);
                            g_webviewController.As(&g_webviewController2);
                            g_webviewController->put_IsVisible(TRUE);
                            UpdateWebViewWindowHandle();
                            UpdateWebViewInputState();
                            UpdateWebViewInputTimer();
                            UpdateWebViewBounds();
                            if (g_webview && !g_pendingHtmlContent.empty())
                            {
                                g_webview->NavigateToString(g_pendingHtmlContent.c_str());
                                g_pendingHtmlContent.clear();
                            }
                            return S_OK;
                        }).Get());
            }).Get());

    return SUCCEEDED(hr);
}

void CloseWebView()
{
    if (g_webviewController)
    {
        g_webviewController->Close();
    }
    if (g_webviewInputTimerActive && g_hwnd)
    {
        KillTimer(g_hwnd, kWebViewInputTimerId);
        g_webviewInputTimerActive = false;
    }
    g_webviewController.Reset();
    g_webviewController2.Reset();
    g_webview.Reset();
    g_webviewWindow = nullptr;
    if (g_webviewLoader)
    {
        FreeLibrary(g_webviewLoader);
        g_webviewLoader = nullptr;
    }
}

void UpdateCustomColorBrush()
{
    if (!g_renderTarget)
    {
        return;
    }
    if (g_customColorBrush)
    {
        g_customColorBrush->Release();
        g_customColorBrush = nullptr;
    }
    D2D1_COLOR_F color = D2D1::ColorF(
        GetRValue(g_customColor) / 255.0f,
        GetGValue(g_customColor) / 255.0f,
        GetBValue(g_customColor) / 255.0f
    );
    g_renderTarget->CreateSolidColorBrush(color, &g_customColorBrush);
}

void UpdateTextFormat()
{
    if (!g_dwriteFactory)
    {
        return;
    }
    if (g_textFormat)
    {
        g_textFormat->Release();
        g_textFormat = nullptr;
    }
    g_dwriteFactory->CreateTextFormat(
        g_textFontName.c_str(),
        nullptr,
        DWRITE_FONT_WEIGHT_REGULAR,
        DWRITE_FONT_STYLE_NORMAL,
        DWRITE_FONT_STRETCH_NORMAL,
        g_textFontSize,
        L"ja-jp",
        &g_textFormat
    );
    if (g_textFormat)
    {
        g_textFormat->SetWordWrapping(g_textWrap ? DWRITE_WORD_WRAPPING_WRAP : DWRITE_WORD_WRAPPING_NO_WRAP);
    }
}

void UpdateTextBrush()
{
    if (!g_renderTarget)
    {
        return;
    }
    if (g_textBrush)
    {
        g_textBrush->Release();
        g_textBrush = nullptr;
    }
    D2D1_COLOR_F color = D2D1::ColorF(
        GetRValue(g_textColor) / 255.0f,
        GetGValue(g_textColor) / 255.0f,
        GetBValue(g_textColor) / 255.0f
    );
    g_renderTarget->CreateSolidColorBrush(color, &g_textBrush);
}

void ResizeWindowByFactor(HWND hwnd, float factor)
{
    RECT rc{};
    GetClientRect(hwnd, &rc);
    float width = static_cast<float>(rc.right - rc.left) * factor;
    float height = static_cast<float>(rc.bottom - rc.top) * factor;
    width = std::max(200.0f, width);
    height = std::max(200.0f, height);
    SetWindowPos(hwnd, nullptr, 0, 0, static_cast<int>(width), static_cast<int>(height),
        SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    InvalidateRect(hwnd, nullptr, TRUE);
}

void ScrollTextBy(float delta)
{
    if (!g_renderTarget || !g_textFormat)
    {
        return;
    }

    D2D1_SIZE_F rtSize = g_renderTarget->GetSize();
    if (rtSize.width <= 0.0f || rtSize.height <= 0.0f)
    {
        return;
    }

    IDWriteTextLayout* layout = nullptr;
    g_dwriteFactory->CreateTextLayout(
        g_textContent.c_str(),
        static_cast<UINT32>(g_textContent.size()),
        g_textFormat,
        rtSize.width - 16.0f,
        10000.0f,
        &layout
    );
    if (!layout)
    {
        return;
    }

    DWRITE_TEXT_METRICS metrics{};
    layout->GetMetrics(&metrics);
    layout->Release();

    float maxScroll = std::max(0.0f, metrics.height + 16.0f - rtSize.height);
    g_textScroll = std::max(0.0f, std::min(g_textScroll + delta, maxScroll));
}

namespace
{
    constexpr int kIdTransparencySelect = 2001;
    constexpr int kIdColor = 2004;
    constexpr int kIdFont = 2005;
    constexpr int kIdFontColor = 2006;
    constexpr int kIdBackColor = 2007;
    constexpr int kIdWrap = 2008;
    constexpr int kIdKeyNext = 2101;
    constexpr int kIdKeyPrev = 2102;
    constexpr int kIdKeyZoomIn = 2103;
    constexpr int kIdKeyZoomOut = 2104;
    constexpr int kIdKeyOpen = 2105;
    constexpr int kIdKeyExit = 2106;
    constexpr int kIdKeyAlwaysOnTop = 2107;
}

enum class PreferredAppMode
{
    Default,
    AllowDark,
    ForceDark,
    ForceLight,
    Max
};

using SetPreferredAppModeFn = PreferredAppMode(WINAPI*)(PreferredAppMode appMode);
using FlushMenuThemesFn = void (WINAPI*)();

SetPreferredAppModeFn g_setPreferredAppMode = nullptr;
FlushMenuThemesFn g_flushMenuThemes = nullptr;

bool IsDarkModeEnabled()
{
    DWORD value = 1;
    DWORD valueSize = sizeof(value);
    if (RegGetValueW(
        HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
        L"AppsUseLightTheme",
        RRF_RT_REG_DWORD,
        nullptr,
        &value,
        &valueSize
    ) == ERROR_SUCCESS)
    {
        return value == 0;
    }
    return false;
}

void InitializeThemeMode()
{
    HMODULE themeModule = GetModuleHandleW(L"uxtheme.dll");
    if (!themeModule)
    {
        return;
    }
    g_setPreferredAppMode = reinterpret_cast<SetPreferredAppModeFn>(
        GetProcAddress(themeModule, MAKEINTRESOURCEA(135))
    );
    g_flushMenuThemes = reinterpret_cast<FlushMenuThemesFn>(
        GetProcAddress(themeModule, MAKEINTRESOURCEA(136))
    );
    if (g_setPreferredAppMode)
    {
        g_setPreferredAppMode(PreferredAppMode::AllowDark);
        if (g_flushMenuThemes)
        {
            g_flushMenuThemes();
        }
    }
}

void ApplyImmersiveDarkMode(HWND target, bool enabled)
{
    const DWORD darkModeAttribute = 20;
    const DWORD darkModeAttributeFallback = 19;
    BOOL useDark = enabled ? TRUE : FALSE;
    DwmSetWindowAttribute(target, darkModeAttribute, &useDark, sizeof(useDark));
    DwmSetWindowAttribute(target, darkModeAttributeFallback, &useDark, sizeof(useDark));
}

void RefreshMenuTheme()
{
    if (g_flushMenuThemes && IsDarkModeEnabled())
    {
        g_flushMenuThemes();
    }
}

static void ApplyExplorerTheme(HWND target)
{
    bool darkMode = IsDarkModeEnabled();
    const wchar_t* themeName = darkMode ? L"DarkMode_Explorer" : L"Explorer";
    ApplyImmersiveDarkMode(target, darkMode);
    SetWindowTheme(target, themeName, nullptr);
    EnumChildWindows(
        target,
        [](HWND child, LPARAM param) -> BOOL
        {
            const wchar_t* theme = reinterpret_cast<const wchar_t*>(param);
            SetWindowTheme(child, theme, nullptr);
            return TRUE;
        },
        reinterpret_cast<LPARAM>(themeName)
    );
}

void ShowSettingsDialog(HWND hwnd)
{

    auto alignDword = [](std::vector<BYTE>& buffer)
    {
        while (buffer.size() % 4 != 0)
        {
            buffer.push_back(0);
        }
    };

    auto appendWord = [](std::vector<BYTE>& buffer, WORD value)
    {
        buffer.push_back(static_cast<BYTE>(value & 0xFF));
        buffer.push_back(static_cast<BYTE>((value >> 8) & 0xFF));
    };

    auto appendDword = [&](std::vector<BYTE>& buffer, DWORD value)
    {
        appendWord(buffer, static_cast<WORD>(value & 0xFFFF));
        appendWord(buffer, static_cast<WORD>((value >> 16) & 0xFFFF));
    };

    auto appendString = [&](std::vector<BYTE>& buffer, const wchar_t* text)
    {
        while (*text)
        {
            appendWord(buffer, static_cast<WORD>(*text));
            ++text;
        }
        appendWord(buffer, 0);
    };

    auto addControl = [&](std::vector<BYTE>& buffer, DWORD style, short x, short y, short cx, short cy, WORD id, WORD classAtom, const wchar_t* text)
    {
        alignDword(buffer);
        appendDword(buffer, style);
        appendDword(buffer, 0);
        appendWord(buffer, static_cast<WORD>(x));
        appendWord(buffer, static_cast<WORD>(y));
        appendWord(buffer, static_cast<WORD>(cx));
        appendWord(buffer, static_cast<WORD>(cy));
        appendWord(buffer, id);
        appendWord(buffer, 0xFFFF);
        appendWord(buffer, classAtom);
        appendString(buffer, text);
        appendWord(buffer, 0);
    };
    auto addControlWithClassName = [&](std::vector<BYTE>& buffer, DWORD style, short x, short y, short cx, short cy, WORD id, const wchar_t* className, const wchar_t* text)
    {
        alignDword(buffer);
        appendDword(buffer, style);
        appendDword(buffer, 0);
        appendWord(buffer, static_cast<WORD>(x));
        appendWord(buffer, static_cast<WORD>(y));
        appendWord(buffer, static_cast<WORD>(cx));
        appendWord(buffer, static_cast<WORD>(cy));
        appendWord(buffer, id);
        appendString(buffer, className);
        appendString(buffer, text);
        appendWord(buffer, 0);
    };

    std::vector<BYTE> tmpl;
    tmpl.reserve(1024);

    constexpr float kDialogScale = 0.9f;
    auto scale = [=](short value)
    {
        return static_cast<short>(std::lround(value * kDialogScale));
    };

    DWORD dialogStyle = WS_POPUP | WS_CAPTION | WS_SYSMENU | DS_MODALFRAME | DS_SETFONT | DS_SHELLFONT;
    appendDword(tmpl, dialogStyle);
    appendDword(tmpl, 0);
    appendWord(tmpl, 25);
    appendWord(tmpl, scale(10));
    appendWord(tmpl, scale(10));
    appendWord(tmpl, scale(248));
    appendWord(tmpl, scale(320));
    appendWord(tmpl, 0);
    appendWord(tmpl, 0);
    appendString(tmpl, L"Settings");
    appendWord(tmpl, static_cast<WORD>(std::lround(10.0f * kDialogScale)));
    appendString(tmpl, L"Segoe UI");

    addControl(tmpl, WS_CHILD | WS_VISIBLE | BS_GROUPBOX, scale(8), scale(6), scale(232), scale(78), 0xFFFF, 0x0080, L"Background of transparent images");
    addControl(tmpl, WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST, scale(16), scale(22), scale(214), scale(80), kIdTransparencySelect, 0x0085, L"");
    addControl(tmpl, WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, scale(16), scale(60), scale(90), scale(16), kIdColor, 0x0080, L"Color...");

    addControl(tmpl, WS_CHILD | WS_VISIBLE | BS_GROUPBOX, scale(8), scale(90), scale(232), scale(78), 0xFFFF, 0x0080, L"Text");
    addControl(tmpl, WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, scale(16), scale(106), scale(90), scale(16), kIdFont, 0x0080, L"Font...");
    addControl(tmpl, WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, scale(112), scale(106), scale(90), scale(16), kIdFontColor, 0x0080, L"Font color");
    addControl(tmpl, WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, scale(112), scale(126), scale(90), scale(16), kIdBackColor, 0x0080, L"Background");
    addControl(tmpl, WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, scale(16), scale(128), scale(80), scale(12), kIdWrap, 0x0080, L"Wrap");

    addControl(tmpl, WS_CHILD | WS_VISIBLE | BS_GROUPBOX, scale(8), scale(174), scale(232), scale(124), 0xFFFF, 0x0080, L"Key Config");
    addControl(tmpl, WS_CHILD | WS_VISIBLE, scale(16), scale(190), scale(110), scale(12), 0xFFFF, 0x0082, L"Next file");
    addControlWithClassName(tmpl, WS_CHILD | WS_VISIBLE | WS_TABSTOP, scale(140), scale(188), scale(88), scale(16), kIdKeyNext, L"msctls_hotkey32", L"");
    addControl(tmpl, WS_CHILD | WS_VISIBLE, scale(16), scale(206), scale(110), scale(12), 0xFFFF, 0x0082, L"Previous file");
    addControlWithClassName(tmpl, WS_CHILD | WS_VISIBLE | WS_TABSTOP, scale(140), scale(204), scale(88), scale(16), kIdKeyPrev, L"msctls_hotkey32", L"");
    addControl(tmpl, WS_CHILD | WS_VISIBLE, scale(16), scale(222), scale(110), scale(12), 0xFFFF, 0x0082, L"Zoom in");
    addControlWithClassName(tmpl, WS_CHILD | WS_VISIBLE | WS_TABSTOP, scale(140), scale(220), scale(88), scale(16), kIdKeyZoomIn, L"msctls_hotkey32", L"");
    addControl(tmpl, WS_CHILD | WS_VISIBLE, scale(16), scale(238), scale(110), scale(12), 0xFFFF, 0x0082, L"Zoom out");
    addControlWithClassName(tmpl, WS_CHILD | WS_VISIBLE | WS_TABSTOP, scale(140), scale(236), scale(88), scale(16), kIdKeyZoomOut, L"msctls_hotkey32", L"");
    addControl(tmpl, WS_CHILD | WS_VISIBLE, scale(16), scale(254), scale(110), scale(12), 0xFFFF, 0x0082, L"Open file");
    addControlWithClassName(tmpl, WS_CHILD | WS_VISIBLE | WS_TABSTOP, scale(140), scale(252), scale(88), scale(16), kIdKeyOpen, L"msctls_hotkey32", L"");
    addControl(tmpl, WS_CHILD | WS_VISIBLE, scale(16), scale(270), scale(110), scale(12), 0xFFFF, 0x0082, L"Exit");
    addControlWithClassName(tmpl, WS_CHILD | WS_VISIBLE | WS_TABSTOP, scale(140), scale(268), scale(88), scale(16), kIdKeyExit, L"msctls_hotkey32", L"");
    addControl(tmpl, WS_CHILD | WS_VISIBLE, scale(16), scale(286), scale(110), scale(12), 0xFFFF, 0x0082, L"Always on Top");
    addControlWithClassName(tmpl, WS_CHILD | WS_VISIBLE | WS_TABSTOP, scale(140), scale(284), scale(88), scale(16), kIdKeyAlwaysOnTop, L"msctls_hotkey32", L"");

    addControl(tmpl, WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, scale(124), scale(304), scale(54), scale(18), IDOK, 0x0080, L"Save");
    addControl(tmpl, WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, scale(184), scale(304), scale(54), scale(18), IDCANCEL, 0x0080, L"Cancel");

    struct DialogState
    {
        TransparencyMode mode;
        COLORREF color;
        std::wstring fontName;
        float fontSize;
        COLORREF fontColor;
        COLORREF backgroundColor;
        bool wrap;
        WORD keyNext;
        WORD keyPrev;
        WORD keyZoomIn;
        WORD keyZoomOut;
        WORD keyOpen;
        WORD keyExit;
        WORD keyAlwaysOnTop;
    } state{ g_transparencyMode, g_customColor, g_textFontName, g_textFontSize, g_textColor, g_textBackground, g_textWrap,
        g_keyNextFile, g_keyPrevFile, g_keyZoomIn, g_keyZoomOut, g_keyOpenFile, g_keyExit, g_keyAlwaysOnTop };

    auto dialogProc = [](HWND dlg, UINT msg, WPARAM wParam, LPARAM lParam) -> INT_PTR
    {
        auto* dialogState = reinterpret_cast<DialogState*>(GetWindowLongPtr(dlg, GWLP_USERDATA));
        switch (msg)
        {
        case WM_INITDIALOG:
        {
            SetWindowLongPtr(dlg, GWLP_USERDATA, lParam);
            dialogState = reinterpret_cast<DialogState*>(lParam);
            ApplyExplorerTheme(dlg);
            SendDlgItemMessage(dlg, kIdTransparencySelect, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Transparent (show desktop)"));
            SendDlgItemMessage(dlg, kIdTransparencySelect, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Checkerboard"));
            SendDlgItemMessage(dlg, kIdTransparencySelect, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Solid color"));
            int selection = (dialogState->mode == TransparencyMode::Transparent) ? 0
                : (dialogState->mode == TransparencyMode::Checkerboard) ? 1 : 2;
            SendDlgItemMessage(dlg, kIdTransparencySelect, CB_SETCURSEL, selection, 0);
            EnableWindow(GetDlgItem(dlg, kIdColor), dialogState->mode == TransparencyMode::SolidColor);
            CheckDlgButton(dlg, kIdWrap, dialogState->wrap ? BST_CHECKED : BST_UNCHECKED);
            SendDlgItemMessage(dlg, kIdKeyNext, HKM_SETHOTKEY, MAKEWORD(dialogState->keyNext, 0), 0);
            SendDlgItemMessage(dlg, kIdKeyPrev, HKM_SETHOTKEY, MAKEWORD(dialogState->keyPrev, 0), 0);
            SendDlgItemMessage(dlg, kIdKeyZoomIn, HKM_SETHOTKEY, MAKEWORD(dialogState->keyZoomIn, 0), 0);
            SendDlgItemMessage(dlg, kIdKeyZoomOut, HKM_SETHOTKEY, MAKEWORD(dialogState->keyZoomOut, 0), 0);
            SendDlgItemMessage(dlg, kIdKeyOpen, HKM_SETHOTKEY, MAKEWORD(dialogState->keyOpen, 0), 0);
            SendDlgItemMessage(dlg, kIdKeyExit, HKM_SETHOTKEY, MAKEWORD(dialogState->keyExit, 0), 0);
            SendDlgItemMessage(dlg, kIdKeyAlwaysOnTop, HKM_SETHOTKEY, MAKEWORD(dialogState->keyAlwaysOnTop, 0), 0);
            return TRUE;
        }
        case WM_COMMAND:
        {
            int id = LOWORD(wParam);
            if (id == kIdTransparencySelect && HIWORD(wParam) == CBN_SELCHANGE)
            {
                int selection = static_cast<int>(SendDlgItemMessage(dlg, kIdTransparencySelect, CB_GETCURSEL, 0, 0));
                dialogState->mode = (selection == 0) ? TransparencyMode::Transparent
                    : (selection == 1) ? TransparencyMode::Checkerboard : TransparencyMode::SolidColor;
                EnableWindow(GetDlgItem(dlg, kIdColor), dialogState->mode == TransparencyMode::SolidColor);
                return TRUE;
            }
            if (id == kIdColor || id == kIdFontColor || id == kIdBackColor)
            {
                CHOOSECOLOR cc{};
                COLORREF custom[16]{};
                cc.lStructSize = sizeof(cc);
                cc.hwndOwner = dlg;
                cc.rgbResult = (id == kIdColor) ? dialogState->color
                    : (id == kIdFontColor) ? dialogState->fontColor : dialogState->backgroundColor;
                cc.lpCustColors = custom;
                cc.Flags = CC_FULLOPEN | CC_RGBINIT;
                if (ChooseColor(&cc))
                {
                    if (id == kIdColor)
                    {
                        dialogState->color = cc.rgbResult;
                    }
                    else if (id == kIdFontColor)
                    {
                        dialogState->fontColor = cc.rgbResult;
                    }
                    else
                    {
                        dialogState->backgroundColor = cc.rgbResult;
                    }
                }
                return TRUE;
            }
            if (id == kIdFont)
            {
                LOGFONT lf{};
                wcsncpy_s(lf.lfFaceName, dialogState->fontName.c_str(), LF_FACESIZE - 1);
                lf.lfHeight = -static_cast<LONG>(dialogState->fontSize);
                CHOOSEFONT cf{};
                cf.lStructSize = sizeof(cf);
                cf.hwndOwner = dlg;
                cf.lpLogFont = &lf;
                cf.Flags = CF_SCREENFONTS | CF_INITTOLOGFONTSTRUCT;
                if (ChooseFont(&cf))
                {
                    dialogState->fontName = lf.lfFaceName;
                    dialogState->fontSize = static_cast<float>(std::abs(lf.lfHeight));
                }
                return TRUE;
            }
            if (id == kIdWrap)
            {
                dialogState->wrap = (IsDlgButtonChecked(dlg, kIdWrap) == BST_CHECKED);
                return TRUE;
            }
            if (id == IDOK)
            {
                auto readHotKey = [&](int controlId, WORD fallback)
                {
                    DWORD value = static_cast<DWORD>(SendDlgItemMessage(dlg, controlId, HKM_GETHOTKEY, 0, 0));
                    WORD key = LOBYTE(value);
                    return key != 0 ? key : fallback;
                };
                int selection = static_cast<int>(SendDlgItemMessage(dlg, kIdTransparencySelect, CB_GETCURSEL, 0, 0));
                dialogState->mode = (selection == 0) ? TransparencyMode::Transparent
                    : (selection == 1) ? TransparencyMode::Checkerboard : TransparencyMode::SolidColor;
                dialogState->wrap = (IsDlgButtonChecked(dlg, kIdWrap) == BST_CHECKED);
                dialogState->keyNext = readHotKey(kIdKeyNext, dialogState->keyNext);
                dialogState->keyPrev = readHotKey(kIdKeyPrev, dialogState->keyPrev);
                dialogState->keyZoomIn = readHotKey(kIdKeyZoomIn, dialogState->keyZoomIn);
                dialogState->keyZoomOut = readHotKey(kIdKeyZoomOut, dialogState->keyZoomOut);
                dialogState->keyOpen = readHotKey(kIdKeyOpen, dialogState->keyOpen);
                dialogState->keyExit = readHotKey(kIdKeyExit, dialogState->keyExit);
                dialogState->keyAlwaysOnTop = readHotKey(kIdKeyAlwaysOnTop, dialogState->keyAlwaysOnTop);
                EndDialog(dlg, IDOK);
                return TRUE;
            }
            if (id == IDCANCEL)
            {
                EndDialog(dlg, IDCANCEL);
                return TRUE;
            }
            break;
        }
        }
        return FALSE;
    };

    INT_PTR result = DialogBoxIndirectParam(
        GetModuleHandle(nullptr),
        reinterpret_cast<DLGTEMPLATE*>(tmpl.data()),
        hwnd,
        dialogProc,
        reinterpret_cast<LPARAM>(&state)
    );

    if (result == IDOK)
    {
        g_transparencyMode = state.mode;
        g_customColor = state.color;
        g_textFontName = state.fontName;
        g_textFontSize = state.fontSize;
        g_textColor = state.fontColor;
        g_textBackground = state.backgroundColor;
        g_textWrap = state.wrap;
        g_keyNextFile = state.keyNext;
        g_keyPrevFile = state.keyPrev;
        g_keyZoomIn = state.keyZoomIn;
        g_keyZoomOut = state.keyZoomOut;
        g_keyOpenFile = state.keyOpen;
        g_keyExit = state.keyExit;
        g_keyAlwaysOnTop = state.keyAlwaysOnTop;
        SaveSettings();
        ApplyTransparencyMode();
        UpdateTextFormat();
        UpdateTextBrush();
        UpdateWebViewInputState();
        InvalidateRect(hwnd, nullptr, TRUE);
    }
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
        bool shouldInclude = g_sortImageOnly ? IsImageFile(filePath) : IsSupportedFile(filePath);
        if (!shouldInclude)
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
    bool result = false;
    if (IsMarkdownFile(g_currentImagePath))
    {
        result = LoadMarkdownFromFile(g_currentImagePath.c_str());
    }
    else if (IsHtmlFile(g_currentImagePath))
    {
        result = LoadHtmlFromFile(g_currentImagePath.c_str());
    }
    else if (IsTextFile(g_currentImagePath))
    {
        result = LoadTextFromFile(g_currentImagePath.c_str());
    }
    else
    {
        result = LoadImageFromFile(g_currentImagePath.c_str());
        if (result)
        {
            UpdateZoomToFitScreen(g_hwnd);
            if (g_hwnd && g_imageHasAlpha && g_transparencyMode == TransparencyMode::Transparent)
            {
                UpdateLayeredWindowFromWic(
                    g_hwnd,
                    g_imageWidth * g_zoom,
                    g_imageHeight * g_zoom
                );
            }
        }
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
    ofn.lpstrFilter = L"Image/Text/HTML/Markdown Files\0*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.webp;*.txt;*.html;*.htm;*.md;*.markdown\0All Files\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;

    if (!GetOpenFileName(&ofn))
    {
        return false;
    }

    if (IsMarkdownFile(filePath))
    {
        if (LoadMarkdownFromFile(filePath))
        {
            RefreshImageList(filePath);
            return true;
        }
    }
    if (IsHtmlFile(filePath))
    {
        if (LoadHtmlFromFile(filePath))
        {
            RefreshImageList(filePath);
            return true;
        }
    }
    if (IsTextFile(filePath))
    {
        if (LoadTextFromFile(filePath))
        {
            RefreshImageList(filePath);
            return true;
        }
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

    wchar_t buffer[128]{};
    auto readKeySetting = [&](const wchar_t* keyName, WORD defaultKey)
    {
        wchar_t keyBuffer[16]{};
        _snwprintf_s(keyBuffer, _TRUNCATE, L"%u", static_cast<unsigned int>(defaultKey));
        GetPrivateProfileStringW(L"KeyConfig", keyName, keyBuffer, buffer, 32, g_iniPath.c_str());
        int value = _wtoi(buffer);
        if (value <= 0 || value > 0xFE)
        {
            return defaultKey;
        }
        return static_cast<WORD>(value);
    };
    GetPrivateProfileStringW(L"Settings", L"SortMode", L"0", buffer, 32, g_iniPath.c_str());
    int sortValue = _wtoi(buffer);
    if (sortValue < 0 || sortValue > 3)
    {
        sortValue = 0;
    }
    g_sortMode = static_cast<SortMode>(sortValue);

    GetPrivateProfileStringW(L"Settings", L"SortImageOnly", L"1", buffer, 32, g_iniPath.c_str());
    g_sortImageOnly = (_wtoi(buffer) != 0);

    GetPrivateProfileStringW(L"Settings", L"AlwaysOnTop", L"0", buffer, 32, g_iniPath.c_str());
    g_alwaysOnTop = (_wtoi(buffer) != 0);

    GetPrivateProfileStringW(L"Settings", L"TransparencyMode", L"0", buffer, 32, g_iniPath.c_str());
    int modeValue = _wtoi(buffer);
    if (modeValue < 0 || modeValue > 2)
    {
        modeValue = 0;
    }
    g_transparencyMode = static_cast<TransparencyMode>(modeValue);

    GetPrivateProfileStringW(L"Settings", L"TransparencyColor", L"0", buffer, 32, g_iniPath.c_str());
    g_customColor = static_cast<COLORREF>(_wtoi(buffer));

    GetPrivateProfileStringW(L"Text", L"FontName", g_textFontName.c_str(), buffer, 32, g_iniPath.c_str());
    if (buffer[0] != L'\0')
    {
        g_textFontName = buffer;
    }
    GetPrivateProfileStringW(L"Text", L"FontSize", L"18", buffer, 32, g_iniPath.c_str());
    g_textFontSize = static_cast<float>(_wtof(buffer));
    if (g_textFontSize < 8.0f)
    {
        g_textFontSize = 8.0f;
    }
    GetPrivateProfileStringW(L"Text", L"FontColor", L"15790320", buffer, 32, g_iniPath.c_str());
    g_textColor = static_cast<COLORREF>(_wtoi(buffer));
    GetPrivateProfileStringW(L"Text", L"BackgroundColor", L"1315860", buffer, 32, g_iniPath.c_str());
    g_textBackground = static_cast<COLORREF>(_wtoi(buffer));
    GetPrivateProfileStringW(L"Text", L"Wrap", L"1", buffer, 32, g_iniPath.c_str());
    g_textWrap = (_wtoi(buffer) != 0);

    g_keyNextFile = readKeySetting(L"NextFile", VK_RIGHT);
    g_keyPrevFile = readKeySetting(L"PrevFile", VK_LEFT);
    g_keyZoomIn = readKeySetting(L"ZoomIn", VK_UP);
    g_keyZoomOut = readKeySetting(L"ZoomOut", VK_DOWN);
    g_keyOpenFile = readKeySetting(L"OpenFile", 'O');
    g_keyExit = readKeySetting(L"Exit", VK_ESCAPE);
    g_keyAlwaysOnTop = readKeySetting(L"AlwaysOnTop", 'P');
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

    _snwprintf_s(buffer, _TRUNCATE, L"%d", g_sortImageOnly ? 1 : 0);
    WritePrivateProfileStringW(L"Settings", L"SortImageOnly", buffer, g_iniPath.c_str());

    _snwprintf_s(buffer, _TRUNCATE, L"%d", g_alwaysOnTop ? 1 : 0);
    WritePrivateProfileStringW(L"Settings", L"AlwaysOnTop", buffer, g_iniPath.c_str());

    _snwprintf_s(buffer, _TRUNCATE, L"%d", static_cast<int>(g_transparencyMode));
    WritePrivateProfileStringW(L"Settings", L"TransparencyMode", buffer, g_iniPath.c_str());

    _snwprintf_s(buffer, _TRUNCATE, L"%u", static_cast<unsigned int>(g_customColor));
    WritePrivateProfileStringW(L"Settings", L"TransparencyColor", buffer, g_iniPath.c_str());

    WritePrivateProfileStringW(L"Text", L"FontName", g_textFontName.c_str(), g_iniPath.c_str());
    _snwprintf_s(buffer, _TRUNCATE, L"%.2f", g_textFontSize);
    WritePrivateProfileStringW(L"Text", L"FontSize", buffer, g_iniPath.c_str());
    _snwprintf_s(buffer, _TRUNCATE, L"%u", static_cast<unsigned int>(g_textColor));
    WritePrivateProfileStringW(L"Text", L"FontColor", buffer, g_iniPath.c_str());
    _snwprintf_s(buffer, _TRUNCATE, L"%u", static_cast<unsigned int>(g_textBackground));
    WritePrivateProfileStringW(L"Text", L"BackgroundColor", buffer, g_iniPath.c_str());
    _snwprintf_s(buffer, _TRUNCATE, L"%d", g_textWrap ? 1 : 0);
    WritePrivateProfileStringW(L"Text", L"Wrap", buffer, g_iniPath.c_str());

    _snwprintf_s(buffer, _TRUNCATE, L"%u", static_cast<unsigned int>(g_keyNextFile));
    WritePrivateProfileStringW(L"KeyConfig", L"NextFile", buffer, g_iniPath.c_str());
    _snwprintf_s(buffer, _TRUNCATE, L"%u", static_cast<unsigned int>(g_keyPrevFile));
    WritePrivateProfileStringW(L"KeyConfig", L"PrevFile", buffer, g_iniPath.c_str());
    _snwprintf_s(buffer, _TRUNCATE, L"%u", static_cast<unsigned int>(g_keyZoomIn));
    WritePrivateProfileStringW(L"KeyConfig", L"ZoomIn", buffer, g_iniPath.c_str());
    _snwprintf_s(buffer, _TRUNCATE, L"%u", static_cast<unsigned int>(g_keyZoomOut));
    WritePrivateProfileStringW(L"KeyConfig", L"ZoomOut", buffer, g_iniPath.c_str());
    _snwprintf_s(buffer, _TRUNCATE, L"%u", static_cast<unsigned int>(g_keyOpenFile));
    WritePrivateProfileStringW(L"KeyConfig", L"OpenFile", buffer, g_iniPath.c_str());
    _snwprintf_s(buffer, _TRUNCATE, L"%u", static_cast<unsigned int>(g_keyExit));
    WritePrivateProfileStringW(L"KeyConfig", L"Exit", buffer, g_iniPath.c_str());
    _snwprintf_s(buffer, _TRUNCATE, L"%u", static_cast<unsigned int>(g_keyAlwaysOnTop));
    WritePrivateProfileStringW(L"KeyConfig", L"AlwaysOnTop", buffer, g_iniPath.c_str());
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
    if (!g_wicSourcePremultiplied || drawWidth <= 0.0f || drawHeight <= 0.0f)
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

    hr = scaler->Initialize(g_wicSourcePremultiplied, width, height, WICBitmapInterpolationModeFant);
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
    if (g_hasHtml)
    {
        UpdateWebViewBounds();
        return;
    }

    if (!g_renderTarget)
    {
        if (!InitDirect2D(hwnd))
        {
            return;
        }
    }

    if (g_hasText)
    {
        g_renderTarget->BeginDraw();
        D2D1_SIZE_F rtSize = g_renderTarget->GetSize();
        if (g_renderTarget && rtSize.width > 0.0f && rtSize.height > 0.0f)
        {
            D2D1_COLOR_F bgColor = D2D1::ColorF(
                GetRValue(g_textBackground) / 255.0f,
                GetGValue(g_textBackground) / 255.0f,
                GetBValue(g_textBackground) / 255.0f
            );
            g_renderTarget->Clear(bgColor);
            if (g_textFormat && g_textBrush)
            {
                IDWriteTextLayout* layout = nullptr;
                g_dwriteFactory->CreateTextLayout(
                    g_textContent.c_str(),
                    static_cast<UINT32>(g_textContent.size()),
                    g_textFormat,
                    rtSize.width - 16.0f,
                    10000.0f,
                    &layout
                );
                if (layout)
                {
                    g_renderTarget->DrawTextLayout(
                        D2D1::Point2F(8.0f, 8.0f - g_textScroll),
                        layout,
                        g_textBrush
                    );
                    layout->Release();
                }
            }
        }
        HRESULT hr = g_renderTarget->EndDraw();
        if (hr == D2DERR_RECREATE_TARGET)
        {
            DiscardRenderTarget();
            InvalidateRect(hwnd, nullptr, TRUE);
        }
    }
    else if (g_bitmap)
    {
        float scale = g_zoom;
        float drawWidth = g_imageWidth * scale;
        float drawHeight = g_imageHeight * scale;

        UpdateWindowSizeToImage(hwnd, drawWidth, drawHeight);

        if (g_imageHasAlpha && g_transparencyMode == TransparencyMode::Transparent)
        {
            UpdateLayeredWindowFromWic(hwnd, drawWidth, drawHeight);
            return;
        }

        g_renderTarget->BeginDraw();
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
        if (g_imageHasAlpha && g_transparencyMode == TransparencyMode::Checkerboard
            && g_checkerBrushA && g_checkerBrushB)
        {
            const float cellSize = 16.0f;
            for (float y = 0.0f; y < drawHeight; y += cellSize)
            {
                for (float x = 0.0f; x < drawWidth; x += cellSize)
                {
                    bool evenCell = (static_cast<int>(x / cellSize) + static_cast<int>(y / cellSize)) % 2 == 0;
                    ID2D1SolidColorBrush* brush = evenCell ? g_checkerBrushA : g_checkerBrushB;
                    g_renderTarget->FillRectangle(
                        D2D1::RectF(x, y, x + cellSize, y + cellSize),
                        brush
                    );
                }
            }
        }
        else if (g_imageHasAlpha && g_transparencyMode == TransparencyMode::SolidColor && g_customColorBrush)
        {
            g_renderTarget->Clear(g_customColorBrush->GetColor());
        }

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

        HRESULT hr = g_renderTarget->EndDraw();
        if (hr == D2DERR_RECREATE_TARGET)
        {
            DiscardRenderTarget();
            InvalidateRect(hwnd, nullptr, TRUE);
        }
    }
    else
    {
        g_renderTarget->BeginDraw();
        g_renderTarget->Clear(D2D1::ColorF(0.121568f, 0.121568f, 0.121568f));

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
        HRESULT hr = g_renderTarget->EndDraw();
        if (hr == D2DERR_RECREATE_TARGET)
        {
            DiscardRenderTarget();
            InvalidateRect(hwnd, nullptr, TRUE);
        }
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
    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_WIN95_CLASSES;
    InitCommonControlsEx(&icc);
    InitializeThemeMode();

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
    ApplyTransparencyMode();
    UpdateTextFormat();
    UpdateTextBrush();
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
        if (IsMarkdownFile(argv[1]))
        {
            loadedImage = LoadMarkdownFromFile(argv[1]);
            if (loadedImage)
            {
                RefreshImageList(argv[1]);
            }
        }
        else if (IsHtmlFile(argv[1]))
        {
            loadedImage = LoadHtmlFromFile(argv[1]);
            if (loadedImage)
            {
                RefreshImageList(argv[1]);
            }
        }
        else if (IsTextFile(argv[1]))
        {
            loadedImage = LoadTextFromFile(argv[1]);
            if (loadedImage)
            {
                RefreshImageList(argv[1]);
            }
        }
        else
        {
            loadedImage = LoadImageFromFile(argv[1]);
            if (loadedImage)
            {
                RefreshImageList(argv[1]);
                UpdateZoomToFitScreen(hwnd);
            }
        }
    }
    if (argv)
    {
        LocalFree(argv);
    }
    if (!loadedImage)
    {
        bool loadedSkin = false;
        std::filesystem::path exeDir;
        wchar_t modulePath[MAX_PATH]{};
        if (GetModuleFileNameW(nullptr, modulePath, MAX_PATH) > 0)
        {
            exeDir = std::filesystem::path(modulePath).parent_path();
        }
        if (!exeDir.empty())
        {
            std::filesystem::path skinPath = exeDir / L"skin.png";
            if (std::filesystem::exists(skinPath))
            {
                loadedSkin = LoadImageFromFile(skinPath.c_str());
                if (loadedSkin)
                {
                    UpdateZoomToFitScreen(hwnd);
                }
            }
        }
        if (loadedSkin)
        {
            UpdateLayeredStyle(g_imageHasAlpha);
        }
        else
        {
            // 仮画像表示（g_bitmap が未設定のため Render でプレースホルダ表示）
        }
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
