#include <windows.h>
#include <dwmapi.h>
#include <gdiplus.h>
#include <shellapi.h>
#include <thread>
#include <string>
#include <vector>
#include <random>
#include <chrono>
#include <locale>
#include <algorithm>
#include <iostream> 
#include <TlHelp32.h> 

// IMPORTANT: Include YOUR VirtualDesktopApi.h here!
#include "VirtualDesktopApi.h" 

#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "comctl32.lib")

// --- Global Variables ---
HWND g_hWnd = nullptr;
HWND g_hMenuWnd = nullptr;
HINSTANCE g_hInstance = nullptr;
NOTIFYICONDATA g_nid = { 0 };
bool g_bPaused = true;
HANDLE g_hPauseEvent = nullptr;
HANDLE g_hStopEvent = nullptr;
std::thread g_logicThread;
std::thread g_debuggerCheckThread;
int g_clients = 0;
int g_desktops = 0;
int g_currentClient = 0;
int g_currentDesktop = 0;
RECT g_stateTextRect = { 0 };

// --- 调整后的 GUI 尺寸参数 ---
// 减小内边距，例如从 15.0f 减到 10.0f
float g_guiPaddingX = 10.0f;
// 减小内边距，例如从 10.0f 减到 7.0f
float g_guiPaddingY = 7.0f;
// 调整信息文本和状态文本之间的间距
float g_stateInfoSpacing = 15.0f;
// 调整字体大小，例如从 12.0f 减到 10.0f 或 11.0f
const float GUI_FONT_SIZE = 10.0f; // 新增常量，方便统一修改字体大小

std::random_device rd;
std::mt19937 gen(rd());
IVirtualDesktopManagerInternal* g_pVDManagerInternal = nullptr;

std::wstring g_randomAppName = L"";

// Unique name for the mutex (IMPORTANT: Replace with YOUR OWN generated GUID!)
// You can use Visual Studio's "Tools" -> "Create GUID" to generate a new one.
// The GUID should be unique to your application.
const wchar_t* const g_szSingleInstanceMutexName = L"Global\\{A1B2C3D4-E5F6-7890-ABCD-EF0123456789}_MushroomFarmMutex";
HANDLE g_hSingleInstanceMutex = nullptr;

// Custom messages
#define WM_UPDATE_GUI  (WM_USER + 1)
#define WM_TRAYICON    (WM_USER + 2)
#define ID_MENU_EXIT   1001

// --- WindowInfo Structure ---
struct WindowInfo {
    HWND hwnd = nullptr;
    std::wstring title;
    RECT clientRectScreenCoords = { 0, 0, 0, 0 };
};

// --- Forward Declarations ---
LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK MenuWindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void LogicThreadProc();
void DebuggerCheckThreadProc();
void PositionWindowAtTop(HWND hWnd);
static float GetDpiScale(HWND hwnd);
void DrawGUI(HDC hdc, RECT rect, float dpiScale);
void DrawMenuGUI(HDC hdc, RECT rect, float dpiScale);
void ShowCustomTrayMenu(HWND hWnd, POINT pt);
std::wstring GenerateRandomFileName(int length);
void SelfRenameExecutable();
void SafeShutdown();

// --- Window Enumeration Callback ---
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam) {
    const int TITLE_SIZE = 256;
    wchar_t windowTitle[TITLE_SIZE];
    GetWindowTextW(hwnd, windowTitle, TITLE_SIZE);
    std::wstring title(windowTitle);

    // Full word match for "Roblox"
    // Check if the title contains "Roblox" AND its length is exactly 6 (for "Roblox")
    if (title.find(L"Roblox") != std::wstring::npos && title.length() == std::wstring(L"Roblox").length()) {
        RECT clientRect;
        if (GetClientRect(hwnd, &clientRect)) {
            POINT topLeft = { clientRect.left, clientRect.top };
            POINT bottomRight = { clientRect.right, clientRect.bottom };
            ClientToScreen(hwnd, &topLeft);
            ClientToScreen(hwnd, &bottomRight);

            WindowInfo info;
            info.hwnd = hwnd;
            info.title = title;
            info.clientRectScreenCoords.left = topLeft.x;
            info.clientRectScreenCoords.top = topLeft.y;
            info.clientRectScreenCoords.right = bottomRight.x;
            info.clientRectScreenCoords.bottom = bottomRight.y;

            std::vector<WindowInfo>* pWindows = reinterpret_cast<std::vector<WindowInfo>*>(lParam);
            pWindows->push_back(info);
        }
    }
    return TRUE;
}

// --- Send Key With Random Delay ---
void SendKeyWithRandomDelay(WORD vk, int minMs = 50, int maxMs = 200, bool preferScanCode = true) {
    INPUT inputDown = {};
    inputDown.type = INPUT_KEYBOARD;
    INPUT inputUp = {};
    inputUp.type = INPUT_KEYBOARD;
    bool useScanCode = preferScanCode;
    UINT scanCodeValue = 0;

    if (preferScanCode) {
        scanCodeValue = MapVirtualKey(vk, MAPVK_VK_TO_VSC);
        if (scanCodeValue == 0) {
            useScanCode = false;
        }
    }

    if (useScanCode) {
        inputDown.ki.wScan = static_cast<WORD>(scanCodeValue);
        inputDown.ki.dwFlags = KEYEVENTF_SCANCODE;
        inputUp.ki.wScan = static_cast<WORD>(scanCodeValue);
        inputUp.ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;
    }
    else {
        inputDown.ki.wVk = vk;
        inputDown.ki.dwFlags = 0;
        inputUp.ki.wVk = vk;
        inputUp.ki.dwFlags = KEYEVENTF_KEYUP;
    }

    SendInput(1, &inputDown, sizeof(INPUT));
    std::uniform_int_distribution<> dis(minMs, maxMs);
    int delay = dis(gen);
    std::this_thread::sleep_for(std::chrono::milliseconds(delay));
    SendInput(1, &inputUp, sizeof(INPUT));
}

// --- Force Foreground Window ---
bool ForceForegroundWindow(HWND hwnd) {
    if (!IsWindow(hwnd)) return false;
    DWORD currentThreadId = GetCurrentThreadId();
    HWND hCurrentForeground = GetForegroundWindow();
    DWORD foregroundThreadId = GetWindowThreadProcessId(hCurrentForeground, NULL);
    bool detachRequired = false;

    if (currentThreadId != foregroundThreadId && hCurrentForeground != hwnd) {
        if (AttachThreadInput(currentThreadId, foregroundThreadId, TRUE)) {
            detachRequired = true;
        }
    }
    ShowWindow(hwnd, SW_RESTORE);
    SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    SetForegroundWindow(hwnd);
    BringWindowToTop(hwnd);
    SetFocus(hwnd);
    if (detachRequired) {
        AttachThreadInput(currentThreadId, foregroundThreadId, FALSE);
    }
    std::this_thread::sleep_for(std::chrono::milliseconds(250));
    return (GetForegroundWindow() == hwnd);
}

// --- Initialize Virtual Desktop Manager ---
bool InitializeVirtualDesktopManager() {
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) return false;

    IServiceProvider* pServiceProvider = nullptr;
    // Use CLSID_ImmersiveShell from your VirtualDesktopApi.h
    hr = CoCreateInstance(CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER, IID_IServiceProvider, (void**)&pServiceProvider);
    if (FAILED(hr)) {
        CoUninitialize();
        return false;
    }

    // Use CLSID_VirtualDesktopManager_Service and IID_IVirtualDesktopManagerInternal from your VirtualDesktopApi.h
    hr = pServiceProvider->QueryService(CLSID_VirtualDesktopManager_Service, IID_IVirtualDesktopManagerInternal, (void**)&g_pVDManagerInternal);
    if (FAILED(hr)) {
        SafeRelease(&pServiceProvider);
        CoUninitialize();
        return false;
    }
    SafeRelease(&pServiceProvider);
    if (!g_pVDManagerInternal) {
        CoUninitialize();
        return false;
    }
    return true;
}

// --- Shutdown Virtual Desktop Manager ---
void ShutdownVirtualDesktopManager() {
    SafeRelease(&g_pVDManagerInternal);
    CoUninitialize();
}

// --- Switch To Desktop By Index ---
bool SwitchToDesktopByIndex(UINT index) {
    if (!g_pVDManagerInternal) return false;

    IObjectArray* pDesktopsArray = nullptr;
    HRESULT hr = g_pVDManagerInternal->GetDesktops(&pDesktopsArray);
    if (FAILED(hr) || !pDesktopsArray) return false;

    IVirtualDesktop* pDesktopToSwitch = nullptr;
    // Use __uuidof(IVirtualDesktop) to get the correct IID for IVirtualDesktop
    hr = pDesktopsArray->GetAt(index, __uuidof(IVirtualDesktop), (void**)&pDesktopToSwitch);
    if (FAILED(hr) || !pDesktopToSwitch) {
        SafeRelease(&pDesktopsArray);
        return false;
    }

    // Use SwitchDesktop from your VirtualDesktopApi.h
    hr = g_pVDManagerInternal->SwitchDesktop(pDesktopToSwitch);
    SafeRelease(&pDesktopToSwitch);
    SafeRelease(&pDesktopsArray);
    return SUCCEEDED(hr);
}

// --- Generate Random File Name ---
std::wstring GenerateRandomFileName(int length) {
    const std::wstring chars = L"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
    std::uniform_int_distribution<> char_dist(0, static_cast<int>(chars.length() - 1));
    std::wstring result;
    for (int i = 0; i < length; ++i) {
        result += chars[char_dist(gen)];
    }
    return result;
}

// --- Self-Rename Executable ---
void SelfRenameExecutable() {
    wchar_t szPath[MAX_PATH];
    GetModuleFileNameW(NULL, szPath, MAX_PATH);
    std::wstring currentPath = szPath;

    size_t lastSlash = currentPath.find_last_of(L"\\");
    std::wstring folderPath = currentPath.substr(0, lastSlash + 1);

    std::uniform_int_distribution<> len_dist(3, 16);
    g_randomAppName = GenerateRandomFileName(len_dist(gen));
    std::wstring newName = g_randomAppName + L".exe";
    std::wstring newPath = folderPath + newName;

    if (!MoveFileExW(currentPath.c_str(), newPath.c_str(), MOVEFILE_REPLACE_EXISTING)) {
        size_t lastDot = currentPath.find_last_of(L".");
        if (lastDot != std::wstring::npos) {
            g_randomAppName = currentPath.substr(lastSlash + 1, lastDot - (lastSlash + 1));
        }
        else {
            g_randomAppName = currentPath.substr(lastSlash + 1);
        }
    }
}

// --- Safe Shutdown ---
void SafeShutdown() {
    if (g_hStopEvent != nullptr) {
        SetEvent(g_hStopEvent);
    }

    if (g_hPauseEvent != nullptr) {
        CloseHandle(g_hPauseEvent);
        g_hPauseEvent = nullptr;
    }
    if (g_hStopEvent != nullptr) {
        CloseHandle(g_hStopEvent);
        g_hStopEvent = nullptr;
    }

    g_hSingleInstanceMutex = nullptr;

    Shell_NotifyIcon(NIM_DELETE, &g_nid);

    TerminateProcess(GetCurrentProcess(), 0);
}

// --- Debugger Check Thread ---
void DebuggerCheckThreadProc() {
    while (WaitForSingleObject(g_hStopEvent, 0) != WAIT_OBJECT_0) {
        if (IsDebuggerPresent()) {
            SafeShutdown();
        }

        BOOL bRemoteDebuggerPresent = FALSE;
        if (CheckRemoteDebuggerPresent(GetCurrentProcess(), &bRemoteDebuggerPresent) && bRemoteDebuggerPresent) {
            SafeShutdown();
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }
}


// --- WinMain ---
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nShowCmd) {

    // --- Single Instance Check ---
    g_hSingleInstanceMutex = CreateMutex(
        NULL,
        TRUE,
        g_szSingleInstanceMutexName
    );

    if (g_hSingleInstanceMutex == NULL) {
        MessageBoxW(NULL, L"无法创建互斥锁，程序无法启动。", L"Mushroom Farm - 错误", MB_OK | MB_ICONERROR);
        return 1;
    }

    DWORD dwLastError = GetLastError();
    if (dwLastError == ERROR_ALREADY_EXISTS) {
        CloseHandle(g_hSingleInstanceMutex);
        g_hSingleInstanceMutex = nullptr;
        MessageBoxW(NULL, L"Mushroom Farm 已经在运行中。请勿重复运行。", L"Mushroom Farm - 警告", MB_OK | MB_ICONWARNING);
        return 1;
    }

    SelfRenameExecutable();

    g_hInstance = hInstance;

    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);

    Gdiplus::GdiplusStartupInput gdipInput;
    ULONG_PTR gdipToken;
    Gdiplus::GdiplusStartup(&gdipToken, &gdipInput, nullptr);

    WNDCLASSEX wc = { sizeof(wc) };
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.lpszClassName = L"RobloxGUI";
    RegisterClassEx(&wc);

    WNDCLASSEX wcMenu = { sizeof(wcMenu) };
    wcMenu.style = CS_HREDRAW | CS_VREDRAW;
    wcMenu.lpfnWndProc = MenuWindowProc;
    wcMenu.hInstance = hInstance;
    wcMenu.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcMenu.lpszClassName = L"RobloxTrayMenu";
    RegisterClassEx(&wcMenu);

    float dpiScale = GetDpiScale(GetDesktopWindow());
    // 使用 GUI_FONT_SIZE
    Gdiplus::Font tempFont(L"Segoe UI", GUI_FONT_SIZE * dpiScale);
    Gdiplus::Graphics tempGraphics(GetDC(NULL));
    Gdiplus::RectF layoutRect(0, 0, 2000, 100);
    Gdiplus::RectF boundsState, boundsInfo;

    tempGraphics.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);
    tempGraphics.MeasureString(L"暂停中", -1, &tempFont, layoutRect, &boundsState);
    std::wstring maxInfoString = L"客户端: 99 | 桌面: 99 | 当前客户端: 99 | 当前桌面: 99";
    tempGraphics.MeasureString(maxInfoString.c_str(), -1, &tempFont, layoutRect, &boundsInfo);

    // 重新计算窗口大小，使用新的 padding 和 spacing
    int desiredWidth = static_cast<int>(std::ceil((g_guiPaddingX * 2 * dpiScale) + boundsState.Width + (g_stateInfoSpacing * dpiScale) + boundsInfo.Width));
    int desiredHeight = static_cast<int>(std::ceil(max(boundsState.Height, boundsInfo.Height) + (g_guiPaddingY * 2 * dpiScale)));
    if (desiredHeight < 50) desiredHeight = 50; // 确保最小高度

    g_hWnd = CreateWindowEx(
        WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_LAYERED,
        wc.lpszClassName, L"",
        WS_POPUP,
        CW_USEDEFAULT, CW_USEDEFAULT, desiredWidth, desiredHeight,
        nullptr, nullptr, hInstance, nullptr);
    if (!g_hWnd) {
        Gdiplus::GdiplusShutdown(gdipToken);
        if (g_hSingleInstanceMutex != nullptr) {
            CloseHandle(g_hSingleInstanceMutex);
            g_hSingleInstanceMutex = nullptr;
        }
        return 1;
    }

    SetWindowLong(g_hWnd, GWL_STYLE, WS_POPUP);
    SetLayeredWindowAttributes(g_hWnd, RGB(0, 0, 0), 255, LWA_ALPHA | LWA_COLORKEY);
    BOOL dark = static_cast<BOOL>(1);
    DwmSetWindowAttribute(g_hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &dark, sizeof(dark));

    PositionWindowAtTop(g_hWnd);
    ShowWindow(g_hWnd, SW_SHOW);
    UpdateWindow(g_hWnd);

    g_nid = { sizeof(NOTIFYICONDATA), g_hWnd, 1, NIF_ICON | NIF_MESSAGE | NIF_TIP, WM_TRAYICON, LoadIcon(nullptr, IDI_APPLICATION) };

    if (!g_randomAppName.empty()) {
        wcsncpy_s(g_nid.szTip, g_randomAppName.c_str(), _TRUNCATE);
    }
    else {
        wcscpy_s(g_nid.szTip, L"Mushroom Farm");
    }
    Shell_NotifyIcon(NIM_ADD, &g_nid);

    g_hPauseEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
    g_hStopEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);

    g_logicThread = std::thread(LogicThreadProc);
    g_debuggerCheckThread = std::thread(DebuggerCheckThreadProc);

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    if (g_hStopEvent != nullptr) {
        SetEvent(g_hStopEvent);
    }
    if (g_logicThread.joinable()) {
        g_logicThread.join();
    }
    if (g_debuggerCheckThread.joinable()) {
        g_debuggerCheckThread.join();
    }

    if (g_hPauseEvent != nullptr) {
        CloseHandle(g_hPauseEvent);
        g_hPauseEvent = nullptr;
    }
    if (g_hStopEvent != nullptr) {
        CloseHandle(g_hStopEvent);
        g_hStopEvent = nullptr;
    }

    if (g_hSingleInstanceMutex != nullptr) {
        CloseHandle(g_hSingleInstanceMutex);
        g_hSingleInstanceMutex = nullptr;
    }

    Shell_NotifyIcon(NIM_DELETE, &g_nid);
    Gdiplus::GdiplusShutdown(gdipToken);
    return 0;
}

// --- Position Window At Top ---
void PositionWindowAtTop(HWND hWnd) {
    RECT rc;
    GetWindowRect(hWnd, &rc);
    int w = rc.right - rc.left;
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    SetWindowPos(hWnd, HWND_TOPMOST, (screenW - w) / 2, 0, 0, 0, SWP_NOSIZE);
}

// --- Draw GUI (Top square corners, bottom rounded corners) ---
void DrawGUI(HDC hdc, RECT rect, float dpiScale) {
    HDC memDC = CreateCompatibleDC(hdc);
    HBITMAP bmp = CreateCompatibleBitmap(hdc, rect.right, rect.bottom);
    HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, bmp);

    Gdiplus::Graphics g(memDC);
    g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
    g.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);

    Gdiplus::SolidBrush clearBrush(Gdiplus::Color(0, 0, 0, 0));
    g.FillRectangle(&clearBrush, 0, 0, rect.right, rect.bottom);

    // 使用 GUI_FONT_SIZE
    Gdiplus::Font font(L"Segoe UI", GUI_FONT_SIZE * dpiScale);
    Gdiplus::RectF layoutRect(0, 0, 2000, 100);
    Gdiplus::RectF boundsState, boundsInfo;
    std::wstring state = g_bPaused ? L"暂停中" : L"运行中";
    std::wstring info = L"客户端: " + std::to_wstring(g_clients) +
        L" | 桌面: " + std::to_wstring(g_desktops) +
        L" | 当前客户端: " + std::to_wstring(g_currentClient) +
        L" | 当前桌面: " + std::to_wstring(g_currentDesktop);

    g.MeasureString(state.c_str(), -1, &font, layoutRect, &boundsState);
    g.MeasureString(info.c_str(), -1, &font, layoutRect, &boundsInfo);

    // 重新计算 totalWidth 和 totalHeight，使用新的 padding 和 spacing
    float totalWidth = (g_guiPaddingX * dpiScale) + boundsState.Width + (g_stateInfoSpacing * dpiScale) + boundsInfo.Width + (g_guiPaddingX * dpiScale);
    float maxHeight = max(boundsState.Height, boundsInfo.Height);
    float totalHeight = maxHeight + (2 * g_guiPaddingY * dpiScale);

    // Resize window
    SetWindowPos(g_hWnd, NULL, 0, 0, static_cast<int>(totalWidth), static_cast<int>(totalHeight), SWP_NOMOVE | SWP_NOZORDER);
    PositionWindowAtTop(g_hWnd);

    Gdiplus::SolidBrush backgroundBrush(Gdiplus::Color(255, 30, 30, 30));
    Gdiplus::GraphicsPath path;
    float cornerRadius = 10.0f * dpiScale;
    Gdiplus::RectF drawRect(0, 0, totalWidth, totalHeight);

    path.AddLine(drawRect.X, drawRect.Y + cornerRadius, drawRect.X, drawRect.Y);
    path.AddLine(drawRect.X, drawRect.Y, drawRect.Width, drawRect.Y);
    path.AddLine(drawRect.Width, drawRect.Y, drawRect.Width, drawRect.Y + cornerRadius);
    path.AddArc(drawRect.Width - 2 * cornerRadius, drawRect.Height - 2 * cornerRadius, 2 * cornerRadius, 2 * cornerRadius, 0, 90);
    path.AddArc(drawRect.X, drawRect.Height - 2 * cornerRadius, 2 * cornerRadius, 2 * cornerRadius, 90, 90);
    path.CloseFigure();

    g.FillPath(&backgroundBrush, &path);

    Gdiplus::SolidBrush txtBrush(Gdiplus::Color(255, 255, 255, 255));
    float textY = g_guiPaddingY * dpiScale;
    g.DrawString(state.c_str(), -1, &font, Gdiplus::PointF(g_guiPaddingX * dpiScale, textY), &txtBrush);
    g.DrawString(info.c_str(), -1, &font, Gdiplus::PointF(g_guiPaddingX * dpiScale + boundsState.Width + g_stateInfoSpacing * dpiScale, textY), &txtBrush);

    g_stateTextRect = {
        static_cast<LONG>(g_guiPaddingX * dpiScale),
        static_cast<LONG>(textY),
        static_cast<LONG>(g_guiPaddingX * dpiScale + boundsState.Width),
        static_cast<LONG>(textY + boundsState.Height)
    };

    BitBlt(hdc, 0, 0, rect.right, rect.bottom, memDC, 0, 0, SRCCOPY);
    SelectObject(memDC, oldBmp);
    DeleteObject(bmp);
    DeleteDC(memDC);
}

// --- Draw Menu GUI (Four rounded corners) ---
void DrawMenuGUI(HDC hdc, RECT rect, float dpiScale) {
    HDC memDC = CreateCompatibleDC(hdc);
    HBITMAP bmp = CreateCompatibleBitmap(hdc, rect.right, rect.bottom);
    HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, bmp);

    Gdiplus::Graphics g(memDC);
    g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
    g.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);

    Gdiplus::SolidBrush clearBrush(Gdiplus::Color(0, 0, 0, 0));
    g.FillRectangle(&clearBrush, 0, 0, rect.right, rect.bottom);

    // 使用 GUI_FONT_SIZE
    Gdiplus::Font font(L"Segoe UI", GUI_FONT_SIZE * dpiScale);
    Gdiplus::RectF layoutRect(0, 0, 2000, 100);
    Gdiplus::RectF boundsExit;
    std::wstring exitText = L"Exit";
    g.MeasureString(exitText.c_str(), -1, &font, layoutRect, &boundsExit);

    float padding = 10.0f * dpiScale;
    float totalWidth = boundsExit.Width + 2 * padding;
    float totalHeight = boundsExit.Height + 2 * padding;

    Gdiplus::SolidBrush backgroundBrush(Gdiplus::Color(255, 30, 30, 30));
    Gdiplus::GraphicsPath path;
    float radius = totalHeight / 2.0f;
    Gdiplus::RectF drawRect(0, 0, totalWidth, totalHeight);

    path.AddArc(drawRect.X, drawRect.Y, totalHeight, totalHeight, 180, 90);
    path.AddArc(drawRect.Width - totalHeight, drawRect.Y, totalHeight, totalHeight, 270, 90);
    path.AddArc(drawRect.Width - totalHeight, drawRect.Height - totalHeight, totalHeight, totalHeight, 0, 90);
    path.AddArc(drawRect.X, drawRect.Height - totalHeight, totalHeight, totalHeight, 90, 90);
    path.CloseFigure();

    g.FillPath(&backgroundBrush, &path);

    Gdiplus::SolidBrush txtBrush(Gdiplus::Color(255, 255, 255, 255));
    g.DrawString(exitText.c_str(), -1, &font, Gdiplus::PointF(padding, padding), &txtBrush);

    BitBlt(hdc, 0, 0, rect.right, rect.bottom, memDC, 0, 0, SRCCOPY);
    SelectObject(memDC, oldBmp);
    DeleteObject(bmp);
    DeleteDC(memDC);
}

// --- Show Custom Tray Menu ---
void ShowCustomTrayMenu(HWND hWnd, POINT pt) {
    if (g_hMenuWnd) {
        DestroyWindow(g_hMenuWnd);
        g_hMenuWnd = nullptr;
    }

    float dpiScale = GetDpiScale(hWnd);
    // 使用 GUI_FONT_SIZE
    Gdiplus::Font font(L"Segoe UI", GUI_FONT_SIZE * dpiScale);
    Gdiplus::Graphics tempGraphics(GetDC(NULL));
    Gdiplus::RectF layoutRect(0, 0, 2000, 100);
    Gdiplus::RectF boundsExit;
    tempGraphics.MeasureString(L"Exit", -1, &font, layoutRect, &boundsExit);

    float padding = 10.0f * dpiScale;
    int menuWidth = static_cast<int>(boundsExit.Width + 2 * padding);
    int menuHeight = static_cast<int>(boundsExit.Height + 2 * padding);

    g_hMenuWnd = CreateWindowEx(
        WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_LAYERED,
        L"RobloxTrayMenu", L"",
        WS_POPUP,
        pt.x, pt.y, menuWidth, menuHeight,
        hWnd, nullptr, g_hInstance, nullptr);

    if (g_hMenuWnd) {
        SetWindowLong(g_hMenuWnd, GWL_STYLE, WS_POPUP);
        SetLayeredWindowAttributes(g_hMenuWnd, RGB(0, 0, 0), 255, LWA_ALPHA | LWA_COLORKEY);
        BOOL dark = static_cast<BOOL>(1);
        DwmSetWindowAttribute(g_hMenuWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &dark, sizeof(dark));

        ShowWindow(g_hMenuWnd, SW_SHOW);
        UpdateWindow(g_hMenuWnd);

        SetCapture(g_hMenuWnd);
    }
}

// --- Logic Thread ---
void LogicThreadProc() {
    SetConsoleOutputCP(CP_UTF8);
    std::locale::global(std::locale(""));

    if (!InitializeVirtualDesktopManager()) {
        // Handle error, maybe log or display a message
    }

    std::uniform_int_distribution<> keySendDelaySec(50, 70);
    HANDLE h[] = { g_hPauseEvent, g_hStopEvent };
    const DWORD CHECK_INTERVAL_MS = 100;

    while (true) {
        HANDLE valid_handles[2] = { nullptr, nullptr };
        int num_valid_handles = 0;
        if (g_hPauseEvent != nullptr) {
            valid_handles[num_valid_handles++] = g_hPauseEvent;
        }
        if (g_hStopEvent != nullptr) {
            valid_handles[num_valid_handles++] = g_hStopEvent;
        }

        DWORD wait_result = WAIT_TIMEOUT;
        if (num_valid_handles > 0) {
            wait_result = WaitForMultipleObjects(num_valid_handles, valid_handles, FALSE, CHECK_INTERVAL_MS);
        }
        else {
            std::this_thread::sleep_for(std::chrono::milliseconds(CHECK_INTERVAL_MS));
        }

        if (g_hStopEvent != nullptr && wait_result == (WAIT_OBJECT_0 + (num_valid_handles - 1))) break;

        if (g_bPaused) {
            while (true) {
                if (g_hPauseEvent != nullptr && WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) {
                    break;
                }
                if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) {
                    goto end_logic_thread;
                }
                std::this_thread::sleep_for(std::chrono::milliseconds(CHECK_INTERVAL_MS));
            }
        }

        UINT desktopCount = 0;
        if (g_pVDManagerInternal) {
            g_pVDManagerInternal->GetCount(&desktopCount);
        }
        else {
            desktopCount = 1;
        }

        if (desktopCount == 0 && g_pVDManagerInternal) {
            PostMessage(g_hWnd, WM_UPDATE_GUI, (0 & 0xFFFF) | ((0 & 0xFFFF) << 16), (0 & 0xFFFF) | ((0 & 0xFFFF) << 16));
            std::this_thread::sleep_for(std::chrono::seconds(1));
            continue;
        }

        for (UINT i = 0; i < desktopCount; ++i) {
            if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) goto end_logic_thread;
            if (g_bPaused) {
                while (true) {
                    if (g_hPauseEvent != nullptr && WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) {
                        break;
                    }
                    if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) {
                        goto end_logic_thread;
                    }
                    std::this_thread::sleep_for(std::chrono::milliseconds(CHECK_INTERVAL_MS));
                }
            }

            if (g_pVDManagerInternal) {
                if (!SwitchToDesktopByIndex(i)) {
                    continue;
                }
                auto start_sleep_vd = std::chrono::high_resolution_clock::now();
                while (std::chrono::duration<double>(std::chrono::high_resolution_clock::now() - start_sleep_vd).count() < 1.0) {
                    if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) goto end_logic_thread;
                    if (g_bPaused) {
                        while (true) {
                            if (g_hPauseEvent != nullptr && WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) {
                                break;
                            }
                            if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) {
                                goto end_logic_thread;
                            }
                            std::this_thread::sleep_for(std::chrono::milliseconds(CHECK_INTERVAL_MS));
                        }
                    }
                }
            }
            else if (i > 0) {
                break;
            }

            std::vector<WindowInfo> windowsOnThisDesktop;
            EnumWindows(EnumWindowsProc, reinterpret_cast<LPARAM>(&windowsOnThisDesktop));

            PostMessage(g_hWnd, WM_UPDATE_GUI,
                (static_cast<UINT>(windowsOnThisDesktop.size()) & 0xFFFF) | ((static_cast<UINT>(desktopCount) & 0xFFFF) << 16),
                (0 & 0xFFFF) | ((static_cast<UINT>(i + 1) & 0xFFFF) << 16));

            if (!windowsOnThisDesktop.empty()) {
                int currentClientIndex = 0;
                for (const auto& win : windowsOnThisDesktop) {
                    if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) goto end_logic_thread;
                    if (g_bPaused) {
                        while (true) {
                            if (g_hPauseEvent != nullptr && WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) {
                                break;
                            }
                            if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) {
                                goto end_logic_thread;
                            }
                            std::this_thread::sleep_for(std::chrono::milliseconds(CHECK_INTERVAL_MS));
                        }
                    }

                    PostMessage(g_hWnd, WM_UPDATE_GUI,
                        (static_cast<UINT>(windowsOnThisDesktop.size()) & 0xFFFF) | ((static_cast<UINT>(desktopCount) & 0xFFFF) << 16),
                        ((static_cast<UINT>(++currentClientIndex)) & 0xFFFF) | ((static_cast<UINT>(i + 1) & 0xFFFF) << 16));

                    if (ForceForegroundWindow(win.hwnd)) {
                        auto start_sleep_activate = std::chrono::high_resolution_clock::now();
                        while (std::chrono::duration<double>(std::chrono::high_resolution_clock::now() - start_sleep_activate).count() < 0.3) {
                            if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) goto end_logic_thread;
                            if (g_bPaused) {
                                while (true) {
                                    if (g_hPauseEvent != nullptr && WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) {
                                        break;
                                    }
                                    if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) {
                                        goto end_logic_thread;
                                    }
                                    std::this_thread::sleep_for(std::chrono::milliseconds(CHECK_INTERVAL_MS));
                                }
                            }
                        }

                        SendKeyWithRandomDelay(0x45, 100, 300, true);
                        auto start_sleep_key_e = std::chrono::high_resolution_clock::now();
                        while (std::chrono::duration<double>(std::chrono::high_resolution_clock::now() - start_sleep_key_e).count() < 0.1) {
                            if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) goto end_logic_thread;
                            if (g_bPaused) {
                                while (true) {
                                    if (g_hPauseEvent != nullptr && WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) {
                                        break;
                                    }
                                    if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) {
                                        goto end_logic_thread;
                                    }
                                    std::this_thread::sleep_for(std::chrono::milliseconds(CHECK_INTERVAL_MS));
                                }
                            }
                        }

                        SendKeyWithRandomDelay(0x48, 100, 300, true);
                    }
                    auto start_sleep_next_win = std::chrono::high_resolution_clock::now();
                    while (std::chrono::duration<double>(std::chrono::high_resolution_clock::now() - start_sleep_next_win).count() < 0.5) {
                        if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) goto end_logic_thread;
                        if (g_bPaused) {
                            while (true) {
                                if (g_hPauseEvent != nullptr && WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) {
                                    break;
                                }
                                if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) {
                                    goto end_logic_thread;
                                }
                                std::this_thread::sleep_for(std::chrono::milliseconds(CHECK_INTERVAL_MS));
                            }
                        }
                    }
                }
            }
            if (desktopCount == 1 && i == 0) break;
        }

        PostMessage(g_hWnd, WM_UPDATE_GUI,
            (0 & 0xFFFF) | ((static_cast<UINT>(desktopCount) & 0xFFFF) << 16),
            (0 & 0xFFFF) | ((0 & 0xFFFF) << 16));

        double roundDelaySeconds = keySendDelaySec(gen);
        auto start_sleep = std::chrono::high_resolution_clock::now();
        while (std::chrono::duration<double>(std::chrono::high_resolution_clock::now() - start_sleep).count() < roundDelaySeconds) {
            if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) goto end_logic_thread;
            if (g_bPaused) {
                while (true) {
                    if (g_hPauseEvent != nullptr && WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) {
                        break;
                    }
                    if (g_hStopEvent != nullptr && WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) {
                        goto end_logic_thread;
                    }
                    std::this_thread::sleep_for(std::chrono::milliseconds(CHECK_INTERVAL_MS));
                }
            }
        }
    }

end_logic_thread:
    ShutdownVirtualDesktopManager();
}

// --- Window Procedure ---
LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE: {
        float dpiScale = GetDpiScale(hWnd);
        HDC hdc = GetDC(hWnd);
        Gdiplus::Graphics g(hdc);
        g.SetTextRenderingHint(Gdiplus::TextRenderingHintClearTypeGridFit);
        // 使用 GUI_FONT_SIZE
        Gdiplus::Font font(L"Segoe UI", GUI_FONT_SIZE * dpiScale);
        Gdiplus::RectF layoutRect(0, 0, 2000, 100);
        Gdiplus::RectF boundsState;
        g.MeasureString(L"暂停中", -1, &font, layoutRect, &boundsState);

        // g_stateTextRect 应该直接使用全局的 padding 值进行计算
        g_stateTextRect = {
            static_cast<LONG>(g_guiPaddingX * dpiScale),
            static_cast<LONG>(g_guiPaddingY * dpiScale),
            static_cast<LONG>(g_guiPaddingX * dpiScale + boundsState.Width),
            static_cast<LONG>(g_guiPaddingY * dpiScale + boundsState.Height)
        };
        ReleaseDC(hWnd, hdc);
        break;
    }
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        RECT rc; GetClientRect(hWnd, &rc);
        DrawGUI(hdc, rc, GetDpiScale(hWnd));
        EndPaint(hWnd, &ps);
        break;
    }
    case WM_LBUTTONDOWN: {
        int x = LOWORD(lParam), y = HIWORD(lParam);
        if (x >= g_stateTextRect.left && x <= g_stateTextRect.right &&
            y >= g_stateTextRect.top && y <= g_stateTextRect.bottom) {
            g_bPaused = !g_bPaused;
            if (g_bPaused) {
                if (g_hPauseEvent != nullptr) ResetEvent(g_hPauseEvent);
            }
            else {
                if (g_hPauseEvent != nullptr) SetEvent(g_hPauseEvent);
            }
            InvalidateRect(hWnd, nullptr, FALSE);
        }
        break;
    }
    case WM_TRAYICON: {
        if (lParam == WM_RBUTTONDOWN) {
            POINT pt; GetCursorPos(&pt);
            SetForegroundWindow(hWnd);
            ShowCustomTrayMenu(hWnd, pt);
        }
        break;
    }
    case WM_UPDATE_GUI: {
        g_clients = LOWORD(wParam);
        g_desktops = HIWORD(wParam);
        g_currentClient = LOWORD(lParam);
        g_currentDesktop = HIWORD(lParam);
        InvalidateRect(hWnd, nullptr, FALSE);
        break;
    }
    case WM_ERASEBKGND:
        return TRUE;
    case WM_NCHITTEST:
        return HTCLIENT;
    case WM_DESTROY:
        if (g_hMenuWnd) {
            DestroyWindow(g_hMenuWnd);
            g_hMenuWnd = nullptr;
        }
        SafeShutdown();
        break;
    default:
        return DefWindowProc(hWnd, uMsg, wParam, lParam);
    }
    return 0;
}

// --- Menu Window Procedure ---
LRESULT CALLBACK MenuWindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE:
        break;
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        RECT rc; GetClientRect(hWnd, &rc);
        DrawMenuGUI(hdc, rc, GetDpiScale(hWnd));
        EndPaint(hWnd, &ps);
        break;
    }
    case WM_LBUTTONDOWN: {
        int x = LOWORD(lParam), y = HIWORD(lParam);
        RECT rc; GetClientRect(hWnd, &rc);
        if (x >= 0 && x <= rc.right && y >= 0 && y <= rc.bottom) {
            PostMessage(GetParent(hWnd), WM_CLOSE, 0, 0);
        }
        break;
    }
    case WM_LBUTTONUP: {
        POINT pt = { LOWORD(lParam), HIWORD(lParam) };
        ClientToScreen(hWnd, &pt);
        RECT rc; GetWindowRect(hWnd, &rc);
        if (!PtInRect(&rc, pt)) {
            DestroyWindow(hWnd);
            g_hMenuWnd = nullptr;
            ReleaseCapture();
        }
        break;
    }
    case WM_KILLFOCUS: {
        DestroyWindow(hWnd);
        g_hMenuWnd = nullptr;
        ReleaseCapture();
        break;
    }
    case WM_ERASEBKGND:
        return TRUE;
    case WM_NCHITTEST:
        return HTCLIENT;
    case WM_DESTROY:
        g_hMenuWnd = nullptr;
        ReleaseCapture();
        break;
    default:
        return DefWindowProc(hWnd, uMsg, wParam, lParam);
    }
    return 0;
}

// --- Get DPI Scaling Factor ---
static float GetDpiScale(HWND hwnd) {
    UINT dpi = GetDpiForWindow(hwnd);
    return dpi / 96.0f;
}