#include <windows.h>
#include <dwmapi.h>
#include <gdiplus.h>
#include <shellapi.h>
#include <thread>
#include <string>
#include <vector>
#include <iostream>
#include <random>
#include <chrono>
#include <locale> // For std::locale
#include <algorithm> // For std::max

#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "comctl32.lib")

// VirtualDesktopApi.h is assumed to be included and handled.
// Do not declare its contents here as per instructions.
#include "VirtualDesktopApi.h"

// --- Global Variables (GUI and Logic) ---
HWND        g_hWnd = nullptr;
HINSTANCE   g_hInstance = nullptr;
NOTIFYICONDATA g_nid = { 0 };
bool        g_bPaused = true; // Initial state: Paused
HANDLE      g_hPauseEvent = nullptr;
HANDLE      g_hStopEvent = nullptr;
std::thread g_logicThread;

// GUI update variables
int         g_clients = 0;
int         g_desktops = 0;
int         g_currentClient = 0;
int         g_currentDesktop = 0;
RECT        g_stateTextRect = { 0 };
float       g_guiPaddingX = 10.0f; // Padding from left edge
float       g_guiPaddingY = 15.0f; // Padding from top edge, adjusted for vertical centering
float       g_stateInfoSpacing = 80.0f; // Space between state text and info text

// Custom messages
#define WM_UPDATE_GUI  (WM_USER + 1)
#define WM_TRAYICON    (WM_USER + 2)
#define ID_MENU_EXIT   1001

// --- WindowInfo Structure (from multi-window code) ---
struct WindowInfo {
    HWND hwnd;
    std::wstring title;
    RECT clientRectScreenCoords;
};

// --- Global Variables (from multi-window code) ---
std::random_device rd_desktop;
std::mt19937 gen_desktop(rd_desktop());
IVirtualDesktopManagerInternal* g_pVDManagerInternal = nullptr;

// --- Forward Declarations ---
LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void LogicThreadProc();
void CenterWindow(HWND hWnd);
static float GetDpiScale(HWND hwnd);
void DrawGUI(HDC hdc, RECT rect, float dpiScale);

// --- Window Enumeration Callback (from multi-window code) ---
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam) {
    const int TITLE_SIZE = 256;
    wchar_t windowTitle[TITLE_SIZE];
    GetWindowTextW(hwnd, windowTitle, TITLE_SIZE);
    std::wstring title(windowTitle);

    if (title.find(L"Roblox") != std::wstring::npos) {
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

// --- Send Key With Random Delay (from multi-window code) ---
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
            // std::wcerr << L"����: �޷���������� 0x" << std::hex << vk << std::dec
            //     << L" ӳ�䵽ɨ���롣�����˵���ʹ��������롣" << std::endl;
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

    if (SendInput(1, &inputDown, sizeof(INPUT)) == 0) {
        // std::wcerr << L"SendInput (����) ʧ�ܡ��������: " << GetLastError() << std::endl;
    }
    std::uniform_int_distribution<> dis(minMs, maxMs);
    int delay = dis(gen_desktop);
    std::this_thread::sleep_for(std::chrono::milliseconds(delay));
    if (SendInput(1, &inputUp, sizeof(INPUT)) == 0) {
        // std::wcerr << L"SendInput (�ͷ�) ʧ�ܡ��������: " << GetLastError() << std::endl;
    }
}

// --- Send Right Click With Random Delay (from multi-window code) ---
void SendRightClickWithRandomDelay(int minMs = 50, int maxMs = 200) {
    INPUT inputDown = {};
    inputDown.type = INPUT_MOUSE;
    inputDown.mi.dwFlags = MOUSEEVENTF_RIGHTDOWN;
    SendInput(1, &inputDown, sizeof(INPUT));
    std::uniform_int_distribution<> dis(minMs, maxMs);
    int delay = dis(gen_desktop);
    std::this_thread::sleep_for(std::chrono::milliseconds(delay));
    INPUT inputUp = {};
    inputUp.type = INPUT_MOUSE;
    inputUp.mi.dwFlags = MOUSEEVENTF_RIGHTUP;
    SendInput(1, &inputUp, sizeof(INPUT));
}

// --- Force Foreground Window (from multi-window code) ---
bool ForceForegroundWindow(HWND hwnd) {
    if (!IsWindow(hwnd)) {
        return false;
    }
    DWORD currentThreadId = GetCurrentThreadId();
    HWND hCurrentForeground = GetForegroundWindow();
    DWORD foregroundThreadId = GetWindowThreadProcessId(hCurrentForeground, NULL);
    bool detachRequired = false;

    if (currentThreadId != foregroundThreadId && hCurrentForeground != hwnd) {
        if (AttachThreadInput(currentThreadId, foregroundThreadId, TRUE)) {
            detachRequired = true;
        }
        else {
            // std::wcerr << L"����: AttachThreadInput(TRUE) ʧ�ܡ�����: " << GetLastError() << std::endl;
        }
    }
    ShowWindow(hwnd, SW_RESTORE);
    SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    SetForegroundWindow(hwnd);
    BringWindowToTop(hwnd);
    SetFocus(hwnd);
    if (detachRequired) {
        if (!AttachThreadInput(currentThreadId, foregroundThreadId, FALSE)) {
            // std::wcerr << L"����: AttachThreadInput(FALSE) ʧ�ܡ�����: " << GetLastError() << std::endl;
        }
    }
    std::this_thread::sleep_for(std::chrono::milliseconds(250));
    return (GetForegroundWindow() == hwnd);
}

// --- Initialize Virtual Desktop Manager (from multi-window code) ---
bool InitializeVirtualDesktopManager() {
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        // std::wcerr << L"CoInitializeEx ʧ�ܡ�HR: 0x" << std::hex << hr << std::endl;
        return false;
    }

    IServiceProvider* pServiceProvider = nullptr;
    hr = CoCreateInstance(CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER, IID_IServiceProvider, (void**)&pServiceProvider);
    if (FAILED(hr)) {
        // std::wcerr << L"CoCreateInstance(CLSID_ImmersiveShell) ʧ�ܡ�HR: 0x" << std::hex << hr << std::endl;
        CoUninitialize();
        return false;
    }

    hr = pServiceProvider->QueryService(CLSID_VirtualDesktopManager_Service, IID_IVirtualDesktopManagerInternal, (void**)&g_pVDManagerInternal);
    if (FAILED(hr)) {
        const GUID IID_IVirtualDesktopManagerInternal_Old = { 0xF31574D6, 0xB682, 0x4CDC, {0xBD, 0x56, 0x18, 0x27, 0x86, 0x0A, 0xBE, 0xC6} };
        hr = pServiceProvider->QueryService(CLSID_VirtualDesktopManager_Service, IID_IVirtualDesktopManagerInternal_Old, (void**)&g_pVDManagerInternal);
        if (FAILED(hr)) {
            // std::wcerr << L"QueryService(CLSID_VirtualDesktopManager_Service) ʧ�� (�¾�IID������ʧ��)��HR: 0x" << std::hex << hr << std::endl;
            SafeRelease(&pServiceProvider);
            CoUninitialize();
            return false;
        }
        // std::wcout << L"ʹ���˾ɰ� IVirtualDesktopManagerInternal IID��" << std::endl;
    }

    SafeRelease(&pServiceProvider);
    if (!g_pVDManagerInternal) {
        CoUninitialize();
        return false;
    }
    // std::wcout << L"���������������ʼ���ɹ���" << std::endl;
    return true;
}

// --- Shutdown Virtual Desktop Manager (from multi-window code) ---
void ShutdownVirtualDesktopManager() {
    SafeRelease(&g_pVDManagerInternal);
    CoUninitialize();
    // std::wcout << L"��������������ѹرա�" << std::endl;
}

// --- Switch To Desktop By Index (from multi-window code) ---
bool SwitchToDesktopByIndex(UINT index) {
    if (!g_pVDManagerInternal) return false;

    IObjectArray* pDesktopsArray = nullptr;
    HRESULT hr = g_pVDManagerInternal->GetDesktops(&pDesktopsArray);
    if (FAILED(hr) || !pDesktopsArray) {
        // std::wcerr << L"GetDesktops ʧ�ܡ�HR: 0x" << std::hex << hr << std::endl;
        return false;
    }

    IVirtualDesktop* pDesktopToSwitch = nullptr;
    hr = pDesktopsArray->GetAt(index, __uuidof(IVirtualDesktop), (void**)&pDesktopToSwitch);
    if (FAILED(hr) || !pDesktopToSwitch) {
        // std::wcerr << L"GetAt(" << index << L") ʧ�ܡ�HR: 0x" << std::hex << hr << std::endl;
        SafeRelease(&pDesktopsArray);
        return false;
    }

    hr = g_pVDManagerInternal->SwitchDesktop(pDesktopToSwitch);
    if (FAILED(hr)) {
        // std::wcerr << L"SwitchDesktop(" << index << L") ʧ�ܡ�HR: 0x" << std::hex << hr << std::endl;
    }

    SafeRelease(&pDesktopToSwitch);
    SafeRelease(&pDesktopsArray);
    return SUCCEEDED(hr);
}

// --- WinMain (GUI Entry Point) ---
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int) {
    g_hInstance = hInstance;

    // DPI ��֪
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);

    // ���� GDI+
    Gdiplus::GdiplusStartupInput gdipInput;
    ULONG_PTR gdipToken;
    Gdiplus::GdiplusStartup(&gdipToken, &gdipInput, nullptr);

    // ע�ᴰ����
    WNDCLASSEX wc = { sizeof(wc) };
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.lpszClassName = L"RobloxGUI";
    RegisterClassEx(&wc);

    // Determine initial window size based on expected text width
    float dpiScale = GetDpiScale(GetDesktopWindow()); // Get DPI from primary monitor
    Gdiplus::Font tempFont(L"Segoe UI", 12.0f * dpiScale);
    Gdiplus::Graphics tempGraphics(GetDC(NULL)); // Use desktop DC for measurement
    Gdiplus::RectF layoutRect(0, 0, 1000, 100);
    Gdiplus::RectF boundsState, boundsInfo;

    tempGraphics.MeasureString(L"��ͣ��", -1, &tempFont, layoutRect, &boundsState);
    std::wstring maxInfoString = L"�ͻ���: 99 | ����: 99 | ��ǰ�ͻ���: 99 | ��ǰ����: 99"; // Max expected length
    tempGraphics.MeasureString(maxInfoString.c_str(), -1, &tempFont, layoutRect, &boundsInfo);

    int desiredWidth = static_cast<int>((g_guiPaddingX * 2) + boundsState.Width + g_stateInfoSpacing + boundsInfo.Width);
    int desiredHeight = 50; // Keep fixed height as before

    // Create window
    g_hWnd = CreateWindowEx(
        WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
        wc.lpszClassName, L"",
        WS_POPUP,
        CW_USEDEFAULT, CW_USEDEFAULT, desiredWidth, desiredHeight, // Use calculated width
        nullptr, nullptr, hInstance, nullptr);
    if (!g_hWnd) return 1;

    // Deep-set the initial window size
    SetWindowPos(g_hWnd, NULL, 0, 0, desiredWidth, desiredHeight, SWP_NOMOVE | SWP_NOZORDER);


    // ��ɫ����֧��
    BOOL dark = TRUE;
    DwmSetWindowAttribute(g_hWnd,
        DWMWA_USE_IMMERSIVE_DARK_MODE, &dark, sizeof(dark));
    // Բ�ǡ��ޱ߿���չ
    MARGINS m = { 0 };
    DwmExtendFrameIntoClientArea(g_hWnd, &m);

    CenterWindow(g_hWnd);
    ShowWindow(g_hWnd, SW_SHOW);
    UpdateWindow(g_hWnd);

    // ����ͼ��
    g_nid = { sizeof(NOTIFYICONDATA), g_hWnd, 1,
              NIF_ICON | NIF_MESSAGE | NIF_TIP,
              WM_TRAYICON,
              LoadIcon(nullptr, IDI_APPLICATION) };
    wcscpy_s(g_nid.szTip, L"Roblox GUI");
    Shell_NotifyIcon(NIM_ADD, &g_nid);

    // �¼����߳�
    g_hPauseEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr); // Initial state: FALSE (non-signaled, i.e., paused)
    g_hStopEvent = CreateEvent(nullptr, TRUE, FALSE, nullptr);
    g_logicThread = std::thread(LogicThreadProc);

    // ��Ϣѭ��
    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // ����
    SetEvent(g_hStopEvent); // Signal the logic thread to stop
    g_logicThread.join();   // Wait for the logic thread to finish
    CloseHandle(g_hPauseEvent);
    CloseHandle(g_hStopEvent);
    Shell_NotifyIcon(NIM_DELETE, &g_nid);
    Gdiplus::GdiplusShutdown(gdipToken);
    return 0;
}

// --- Center Window Helper ---
void CenterWindow(HWND hWnd) {
    RECT rc;
    GetWindowRect(hWnd, &rc);
    int w = rc.right - rc.left;
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    SetWindowPos(hWnd, nullptr, (screenW - w) / 2, 0, 0, 0,
        SWP_NOSIZE | SWP_NOZORDER);
}

// --- Draw GUI ---
void DrawGUI(HDC hdc, RECT rect, float dpiScale) {
    // ˫����
    HDC   memDC = CreateCompatibleDC(hdc);
    HBITMAP bmp = CreateCompatibleBitmap(hdc, rect.right, rect.bottom);
    HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, bmp);

    // ����
    HBRUSH brush = CreateSolidBrush(RGB(30, 30, 30));
    FillRect(memDC, &rect, brush);
    DeleteObject(brush);

    // GDI+ �����ı�
    Gdiplus::Graphics g(memDC);
    g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
    Gdiplus::Font font(L"Segoe UI", 12.0f * dpiScale); // Use float for font size
    Gdiplus::SolidBrush txtBrush(Gdiplus::Color(255, 255, 255, 255));

    std::wstring state = g_bPaused ? L"��ͣ��" : L"������";
    Gdiplus::RectF boundsState;
    Gdiplus::RectF layoutRect(0, 0, 1000, 100); // Large enough layout for measurement
    g.MeasureString(state.c_str(), -1, &font, layoutRect, &boundsState);

    // Calculate vertical center for text
    float textHeight = boundsState.Height;
    float verticalOffset = (rect.bottom - rect.top - textHeight) / 2.0f;


    g.DrawString(state.c_str(), -1, &font,
        Gdiplus::PointF(g_guiPaddingX * dpiScale, verticalOffset),
        &txtBrush);

    std::wstring info =
        L"�ͻ���: " + std::to_wstring(g_clients) +
        L" | ����: " + std::to_wstring(g_desktops) +
        L" | ��ǰ�ͻ���: " + std::to_wstring(g_currentClient) +
        L" | ��ǰ����: " + std::to_wstring(g_currentDesktop);

    g.DrawString(info.c_str(), -1, &font,
        Gdiplus::PointF((g_guiPaddingX + g_stateInfoSpacing) * dpiScale, verticalOffset),
        &txtBrush);

    BitBlt(hdc, 0, 0, rect.right, rect.bottom, memDC, 0, 0, SRCCOPY);

    // ����
    SelectObject(memDC, oldBmp);
    DeleteObject(bmp);
    DeleteDC(memDC);
}

// --- Logic Thread (Combined) ---
void LogicThreadProc() {
    SetConsoleOutputCP(CP_UTF8);
    std::locale::global(std::locale("")); // Set global locale for wide character output

    if (!InitializeVirtualDesktopManager()) {
        // std::wcerr << L"�޷���ʼ��������������������򽫽��ڵ�ǰ�������У�������ã���" << std::endl;
    }

    std::uniform_real_distribution<> delayDist(5.0, 10.0);

    HANDLE h[] = { g_hPauseEvent, g_hStopEvent };
    const DWORD CHECK_INTERVAL_MS = 100; // Check for stop/pause events every 100ms

    while (true) {
        // Primary wait for pause/stop. If paused, it will wait on g_hPauseEvent, otherwise continue.
        // The timeout ensures we don't get stuck and can always respond to stop.
        DWORD wait_result = WaitForMultipleObjects(2, h, FALSE, CHECK_INTERVAL_MS);

        if (wait_result == WAIT_OBJECT_0 + 1) { // g_hStopEvent signaled
            break; // Exit the loop and terminate the thread
        }
        if (wait_result == WAIT_OBJECT_0) { // g_hPauseEvent signaled (meaning not paused, or resumed)
            // Continue with the logic
        }
        // If wait_result is WAIT_TIMEOUT, or if g_hPauseEvent is signaled, we proceed.
        // If g_hPauseEvent is NOT signaled (i.e., g_bPaused is true), the next check will handle it.

        // If paused, actively wait for resume or stop
        if (g_bPaused) {
            while (WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) != WAIT_OBJECT_0) {
                if (WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) {
                    goto end_logic_thread;
                }
            }
            // If we reach here, it means g_hPauseEvent was signaled (unpaused).
        }


        UINT desktopCount = 0;
        if (g_pVDManagerInternal) {
            g_pVDManagerInternal->GetCount(&desktopCount);
        }
        else {
            desktopCount = 1; // If VD manager not initialized, assume only current desktop
        }

        if (desktopCount == 0 && g_pVDManagerInternal) {
            // std::wcout << L"δ�ҵ��������棨����GetCount����0�����ȴ�������..." << std::endl;
            PostMessage(g_hWnd, WM_UPDATE_GUI, (0 & 0xFFFF) | ((0 & 0xFFFF) << 16), (0 & 0xFFFF) | ((0 & 0xFFFF) << 16));
            std::this_thread::sleep_for(std::chrono::seconds(1)); // Short sleep before retry
            continue;
        }

        // std::wcout << L"��ʼ��һ��ɨ�裬������ " << desktopCount << L" ���������档" << std::endl;

        for (UINT i = 0; i < desktopCount; ++i) {
            // Check for stop/pause at the start of each desktop iteration
            if (WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) { goto end_logic_thread; }
            if (g_bPaused) { while (WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) != WAIT_OBJECT_0) { if (WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) { goto end_logic_thread; } } }

            if (g_pVDManagerInternal) {
                // std::wcout << L"\n--- �����л������� " << (i + 1) << L" ---" << std::endl;
                if (!SwitchToDesktopByIndex(i)) {
                    // std::wcerr << L"�л������� " << (i + 1) << L" ʧ�ܣ����������档" << std::endl;
                    continue;
                }
                // Sleep, but check for stop/pause
                auto start_sleep_vd = std::chrono::high_resolution_clock::now();
                while (std::chrono::duration<double>(std::chrono::high_resolution_clock::now() - start_sleep_vd).count() < 1.0) { // 1000ms
                    if (WaitForSingleObject(g_hStopEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) { goto end_logic_thread; }
                    if (g_bPaused) { while (WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) != WAIT_OBJECT_0) { if (WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) { goto end_logic_thread; } } }
                }

            }
            else if (i > 0) {
                break; // If no VD manager, only process desktop 0
            }

            std::vector<WindowInfo> windowsOnThisDesktop;
            EnumWindows(EnumWindowsProc, reinterpret_cast<LPARAM>(&windowsOnThisDesktop));

            // Update GUI with current desktop and client count
            PostMessage(g_hWnd, WM_UPDATE_GUI,
                (windowsOnThisDesktop.size() & 0xFFFF) | ((desktopCount & 0xFFFF) << 16),
                (0 & 0xFFFF) | ((i + 1 & 0xFFFF) << 16));


            if (windowsOnThisDesktop.empty()) {
                // std::wcout << L"���� " << (i + 1) << L" ��δ�ҵ��κ� Roblox ���ڡ�" << std::endl;
            }
            else {
                // std::wcout << L"������ " << (i + 1) << L" ���ҵ� " << windowsOnThisDesktop.size() << L" �� Roblox ���ڣ���ʼ����..." << std::endl;
                int currentClientIndex = 0;
                for (const auto& win : windowsOnThisDesktop) {
                    // Check for stop/pause at the start of each client iteration
                    if (WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) { goto end_logic_thread; }
                    if (g_bPaused) { while (WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) != WAIT_OBJECT_0) { if (WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) { goto end_logic_thread; } } }

                    // Update GUI with current client being processed
                    PostMessage(g_hWnd, WM_UPDATE_GUI,
                        (windowsOnThisDesktop.size() & 0xFFFF) | ((desktopCount & 0xFFFF) << 16),
                        ((++currentClientIndex) & 0xFFFF) | ((i + 1 & 0xFFFF) << 16));

                    // std::wcout << L"  ׼����������: \"" << win.title << L"\" (λ������ " << (i + 1) << L")" << std::endl;

                    // std::wcout << L"    ���Լ����: \"" << win.title << L"\"..." << std::endl;
                    if (ForceForegroundWindow(win.hwnd)) {
                        // std::wcout << L"    ����: \"" << win.title << L"\" �Ѽ��" << std::endl;

                        // std::wcout << L"    �ȴ�������ʱ (300ms) ��������..." << std::endl;
                        // Sleep, but check for stop/pause
                        auto start_sleep_activate = std::chrono::high_resolution_clock::now();
                        while (std::chrono::duration<double>(std::chrono::high_resolution_clock::now() - start_sleep_activate).count() < 0.3) { // 300ms
                            if (WaitForSingleObject(g_hStopEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) { goto end_logic_thread; }
                            if (g_bPaused) { while (WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) != WAIT_OBJECT_0) { if (WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) { goto end_logic_thread; } } }
                        }


                        // std::wcout << L"    ��ʼģ�ⰴ����������..." << std::endl;

                        // std::wcout << L"      ���� 'E' �� (����ʹ��ɨ����)..." << std::endl;
                        SendKeyWithRandomDelay(0x45, 100, 300, true);
                        // std::wcout << L"      ���²��ͷŰ��� 'E' ��ɡ�" << std::endl;

                        // Sleep, but check for stop/pause
                        auto start_sleep_key = std::chrono::high_resolution_clock::now();
                        while (std::chrono::duration<double>(std::chrono::high_resolution_clock::now() - start_sleep_key).count() < 0.1) { // 100ms
                            if (WaitForSingleObject(g_hStopEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) { goto end_logic_thread; }
                            if (g_bPaused) { while (WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) != WAIT_OBJECT_0) { if (WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) { goto end_logic_thread; } } }
                        }

                        // std::wcout << L"      ��������Ҽ����..." << std::endl;
                        SendRightClickWithRandomDelay(100, 300);
                        // std::wcout << L"      ģ������Ҽ������ɡ�" << std::endl;
                        // std::wcout << L"    ���� \"" << win.title << L"\" ������ɡ�" << std::endl;
                    }
                    else {
                        // std::wcerr << L"    !!! ����: �޷������: \"" << win.title << L"\". ������ǰ���ڲ�����" << std::endl;
                    }
                    // std::wcout << L"  ------------------------------------" << std::endl;
                    // Sleep, but check for stop/pause
                    auto start_sleep_next_win = std::chrono::high_resolution_clock::now();
                    while (std::chrono::duration<double>(std::chrono::high_resolution_clock::now() - start_sleep_next_win).count() < 0.5) { // 500ms
                        if (WaitForSingleObject(g_hStopEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) { goto end_logic_thread; }
                        if (g_bPaused) { while (WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) != WAIT_OBJECT_0) { if (WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) { goto end_logic_thread; } } }
                    }
                }
            }
            // std::wcout << L"--- ���� " << (i + 1) << L" ������� ---" << std::endl;
            if (desktopCount == 1 && i == 0) break; // If only one desktop and processed, break
        }

        // std::wcout << L"\n����ɶ������������������д��ڵ�һ�ֲ�����" << std::endl;
        double roundDelaySeconds = delayDist(gen_desktop);
        // std::wcout << L"��һ������ɨ�迪ʼǰ�ȴ�: " << roundDelaySeconds << L" ��..." << std::endl;

        // Reset GUI info at the end of a full cycle
        PostMessage(g_hWnd, WM_UPDATE_GUI,
            (0 & 0xFFFF) | ((desktopCount & 0xFFFF) << 16),
            (0 & 0xFFFF) | ((0 & 0xFFFF) << 16));

        // Sleep, but also check for stop/pause
        auto start_sleep = std::chrono::high_resolution_clock::now();
        while (std::chrono::duration<double>(std::chrono::high_resolution_clock::now() - start_sleep).count() < roundDelaySeconds) {
            if (WaitForSingleObject(g_hStopEvent, CHECK_INTERVAL_MS) == WAIT_OBJECT_0) { // Check every CHECK_INTERVAL_MS
                goto end_logic_thread;
            }
            if (g_bPaused) { // If paused during sleep
                while (WaitForSingleObject(g_hPauseEvent, CHECK_INTERVAL_MS) != WAIT_OBJECT_0) {
                    if (WaitForSingleObject(g_hStopEvent, 0) == WAIT_OBJECT_0) {
                        goto end_logic_thread;
                    }
                }
            }
        }
        // std::wcout << L"\n===========================================================\n" << std::endl;
    }

end_logic_thread:
    ShutdownVirtualDesktopManager();
}

// --- Window Procedure (GUI Event Handling) ---
LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg,
    WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE: {
        // Measure text for accurate clickable area and window sizing
        float dpiScale = GetDpiScale(hWnd);
        HDC hdc = GetDC(hWnd);
        Gdiplus::Graphics g(hdc);
        Gdiplus::Font font(L"Segoe UI", 12.0f * dpiScale);
        Gdiplus::RectF layoutRect(0, 0, 1000, 100); // Large enough for measurement
        Gdiplus::RectF boundsState, boundsInfo;

        g.MeasureString(L"��ͣ��", -1, &font, layoutRect, &boundsState);
        std::wstring maxInfoString = L"�ͻ���: 99 | ����: 99 | ��ǰ�ͻ���: 99 | ��ǰ����: 99"; // Max expected length for measurement
        g.MeasureString(maxInfoString.c_str(), -1, &font, layoutRect, &boundsInfo);

        // Calculate the clickable rectangle for the state text
        g_stateTextRect = {
            static_cast<LONG>(g_guiPaddingX * dpiScale),
            static_cast<LONG>(g_guiPaddingY * dpiScale), // Top coordinate for the text
            static_cast<LONG>((g_guiPaddingX * dpiScale) + boundsState.Width),
            static_cast<LONG>((g_guiPaddingY * dpiScale) + boundsState.Height)
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
        if (x >= g_stateTextRect.left &&
            x <= g_stateTextRect.right &&
            y >= g_stateTextRect.top &&
            y <= g_stateTextRect.bottom) {
            g_bPaused = !g_bPaused;
            if (g_bPaused) {
                ResetEvent(g_hPauseEvent); // Set to non-signaled (paused)
            }
            else {
                SetEvent(g_hPauseEvent);   // Set to signaled (resumed)
            }
            InvalidateRect(hWnd, nullptr, FALSE); // Redraw GUI
        }
        break;
    }
    case WM_TRAYICON: {
        if (lParam == WM_RBUTTONDOWN) {
            // Owner-draw �˵�
            HMENU hMenu = CreatePopupMenu();
            AppendMenu(hMenu, MF_OWNERDRAW, ID_MENU_EXIT, L"Exit");

            POINT pt; GetCursorPos(&pt);
            SetForegroundWindow(hWnd);
            // ����������ѡ�� ID
            int cmd = TrackPopupMenu(
                hMenu,
                TPM_RETURNCMD | TPM_BOTTOMALIGN | TPM_LEFTALIGN,
                pt.x, pt.y,
                0, hWnd, nullptr);
            DestroyMenu(hMenu);
            if (cmd == ID_MENU_EXIT) {
                // Trigger close process
                PostMessage(hWnd, WM_CLOSE, 0, 0);
            }
        }
        break;
    }
    case WM_MEASUREITEM: {
        LPMEASUREITEMSTRUCT mis = (LPMEASUREITEMSTRUCT)lParam;
        if (mis->CtlType == ODT_MENU && mis->itemID == ID_MENU_EXIT) {
            float s = GetDpiScale(hWnd);
            // Text size
            HDC hdc = GetDC(hWnd);
            LOGFONT lf = {};
            lf.lfHeight = LONG(-12 * s);
            wcscpy_s(lf.lfFaceName, L"Segoe UI");
            HFONT hFont = CreateFontIndirect(&lf);
            HFONT oldF = (HFONT)SelectObject(hdc, hFont);
            SIZE sz; GetTextExtentPoint32(hdc, L"Exit", 4, &sz);
            SelectObject(hdc, oldF);
            DeleteObject(hFont);
            ReleaseDC(hWnd, hdc);
            mis->itemWidth = sz.cx + LONG(20 * s);
            mis->itemHeight = sz.cy + LONG(10 * s);
        }
        return TRUE;
    }
    case WM_DRAWITEM: {
        LPDRAWITEMSTRUCT dis = (LPDRAWITEMSTRUCT)lParam;
        if (dis->CtlType == ODT_MENU && dis->itemID == ID_MENU_EXIT) {
            // Background
            HBRUSH hbr = CreateSolidBrush(RGB(45, 45, 45));
            FillRect(dis->hDC, &dis->rcItem, hbr);
            DeleteObject(hbr);
            // Text
            SetBkMode(dis->hDC, TRANSPARENT);
            SetTextColor(dis->hDC, RGB(255, 255, 255));
            float s = GetDpiScale(hWnd);
            LOGFONT lf = {};
            lf.lfHeight = LONG(-12 * s);
            wcscpy_s(lf.lfFaceName, L"Segoe UI");
            HFONT hFont = CreateFontIndirect(&lf);
            HFONT oldF = (HFONT)SelectObject(dis->hDC, hFont);
            DrawText(dis->hDC, L"Exit", -1, &dis->rcItem,
                DT_SINGLELINE | DT_VCENTER | DT_CENTER);
            SelectObject(dis->hDC, oldF);
            DeleteObject(hFont);
        }
        return TRUE;
    }
    case WM_COMMAND: {
        // Not handling Exit here anymore
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
        return TRUE; // Skip background erase
    case WM_NCHITTEST:
        return HTCLIENT; // Prevent dragging by title bar
    case WM_DESTROY:
        // Stop the logic thread, then quit message loop
        SetEvent(g_hStopEvent);
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, uMsg, wParam, lParam);
    }
    return 0;
}

// Helper: Get DPI scaling factor for a window
static float GetDpiScale(HWND hwnd) {
    UINT dpi = GetDpiForWindow(hwnd);
    return dpi / 96.0f;
}