#include <windows.h>
#include <iostream>
#include <vector>
#include <string>
#include <random>
#include <chrono>
#include <thread>
#include <locale>
#include "VirtualDesktopApi.h"

struct WindowInfo {
    HWND hwnd;
    std::wstring title;
    RECT clientRectScreenCoords;
};

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

std::random_device rd_desktop;
std::mt19937 gen_desktop(rd_desktop());

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
            std::wcerr << L"警告: 无法将虚拟键码 0x" << std::hex << vk << std::dec
                << L" 映射到扫描码。将回退到仅使用虚拟键码。" << std::endl;
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
        std::wcerr << L"SendInput (按下) 失败。错误代码: " << GetLastError() << std::endl;
    }
    std::uniform_int_distribution<> dis(minMs, maxMs);
    int delay = dis(gen_desktop);
    std::this_thread::sleep_for(std::chrono::milliseconds(delay));
    if (SendInput(1, &inputUp, sizeof(INPUT)) == 0) {
        std::wcerr << L"SendInput (释放) 失败。错误代码: " << GetLastError() << std::endl;
    }
}

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
            std::wcerr << L"警告: AttachThreadInput(TRUE) 失败。错误: " << GetLastError() << std::endl;
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
            std::wcerr << L"警告: AttachThreadInput(FALSE) 失败。错误: " << GetLastError() << std::endl;
        }
    }
    std::this_thread::sleep_for(std::chrono::milliseconds(250));
    return (GetForegroundWindow() == hwnd);
}

IVirtualDesktopManagerInternal* g_pVDManagerInternal = nullptr;

bool InitializeVirtualDesktopManager() {
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        std::wcerr << L"CoInitializeEx 失败。HR: 0x" << std::hex << hr << std::endl;
        return false;
    }

    IServiceProvider* pServiceProvider = nullptr;
    hr = CoCreateInstance(CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER, IID_IServiceProvider, (void**)&pServiceProvider);
    if (FAILED(hr)) {
        std::wcerr << L"CoCreateInstance(CLSID_ImmersiveShell) 失败。HR: 0x" << std::hex << hr << std::endl;
        CoUninitialize();
        return false;
    }

    hr = pServiceProvider->QueryService(CLSID_VirtualDesktopManager_Service, IID_IVirtualDesktopManagerInternal, (void**)&g_pVDManagerInternal);
    if (FAILED(hr)) {
        const GUID IID_IVirtualDesktopManagerInternal_Old = { 0xF31574D6, 0xB682, 0x4CDC, {0xBD, 0x56, 0x18, 0x27, 0x86, 0x0A, 0xBE, 0xC6} };
        hr = pServiceProvider->QueryService(CLSID_VirtualDesktopManager_Service, IID_IVirtualDesktopManagerInternal_Old, (void**)&g_pVDManagerInternal);
        if (FAILED(hr)) {
            std::wcerr << L"QueryService(CLSID_VirtualDesktopManager_Service) 失败 (新旧IID均尝试失败)。HR: 0x" << std::hex << hr << std::endl;
            SafeRelease(&pServiceProvider);
            CoUninitialize();
            return false;
        }
        std::wcout << L"使用了旧版 IVirtualDesktopManagerInternal IID。" << std::endl;
    }

    SafeRelease(&pServiceProvider);
    if (!g_pVDManagerInternal) {
        CoUninitialize();
        return false;
    }
    std::wcout << L"虚拟桌面管理器初始化成功。" << std::endl;
    return true;
}

void ShutdownVirtualDesktopManager() {
    SafeRelease(&g_pVDManagerInternal);
    CoUninitialize();
    std::wcout << L"虚拟桌面管理器已关闭。" << std::endl;
}

bool SwitchToDesktopByIndex(UINT index) {
    if (!g_pVDManagerInternal) return false;

    IObjectArray* pDesktopsArray = nullptr;
    HRESULT hr = g_pVDManagerInternal->GetDesktops(&pDesktopsArray);
    if (FAILED(hr) || !pDesktopsArray) {
        std::wcerr << L"GetDesktops 失败。HR: 0x" << std::hex << hr << std::endl;
        return false;
    }

    IVirtualDesktop* pDesktopToSwitch = nullptr;
    hr = pDesktopsArray->GetAt(index, __uuidof(IVirtualDesktop), (void**)&pDesktopToSwitch);
    if (FAILED(hr) || !pDesktopToSwitch) {
        std::wcerr << L"GetAt(" << index << L") 失败。HR: 0x" << std::hex << hr << std::endl;
        SafeRelease(&pDesktopsArray);
        return false;
    }

    hr = g_pVDManagerInternal->SwitchDesktop(pDesktopToSwitch);
    if (FAILED(hr)) {
        std::wcerr << L"SwitchDesktop(" << index << L") 失败。HR: 0x" << std::hex << hr << std::endl;
    }

    SafeRelease(&pDesktopToSwitch);
    SafeRelease(&pDesktopsArray);
    return SUCCEEDED(hr);
}

int main() {
    SetConsoleOutputCP(CP_UTF8);
    std::locale::global(std::locale(""));

    if (!InitializeVirtualDesktopManager()) {
        std::wcerr << L"无法初始化虚拟桌面管理器，程序将仅在当前桌面运行（如果适用）。" << std::endl;
    }

    std::uniform_real_distribution<> delayDist(5.0, 10.0);
    std::wcout << L"程序启动，开始多桌面循环操作..." << std::endl;

    while (true) {
        UINT desktopCount = 0;
        if (g_pVDManagerInternal) {
            g_pVDManagerInternal->GetCount(&desktopCount);
        }
        else {
            desktopCount = 1;
            std::wcout << L"虚拟桌面管理器未初始化，仅扫描当前桌面。" << std::endl;
        }

        if (desktopCount == 0 && g_pVDManagerInternal) {
            std::wcout << L"未找到虚拟桌面（或者GetCount返回0），等待后重试..." << std::endl;
            std::this_thread::sleep_for(std::chrono::seconds(5));
            continue;
        }

        std::wcout << L"开始新一轮扫描，共发现 " << desktopCount << L" 个虚拟桌面。" << std::endl;

        for (UINT i = 0; i < desktopCount; ++i) {
            if (g_pVDManagerInternal) {
                std::wcout << L"\n--- 正在切换到桌面 " << (i + 1) << L" ---" << std::endl;
                if (!SwitchToDesktopByIndex(i)) {
                    std::wcerr << L"切换到桌面 " << (i + 1) << L" 失败，跳过此桌面。" << std::endl;
                    continue;
                }
                std::this_thread::sleep_for(std::chrono::milliseconds(1000));
            }
            else if (i > 0) {
                break;
            }

            std::vector<WindowInfo> windowsOnThisDesktop;
            EnumWindows(EnumWindowsProc, reinterpret_cast<LPARAM>(&windowsOnThisDesktop));

            if (windowsOnThisDesktop.empty()) {
                std::wcout << L"桌面 " << (i + 1) << L" 上未找到任何 Roblox 窗口。" << std::endl;
            }
            else {
                std::wcout << L"在桌面 " << (i + 1) << L" 上找到 " << windowsOnThisDesktop.size() << L" 个 Roblox 窗口，开始处理..." << std::endl;
                for (const auto& win : windowsOnThisDesktop) {
                    std::wcout << L"  准备操作窗口: \"" << win.title << L"\" (位于桌面 " << (i + 1) << L")" << std::endl;

                    std::wcout << L"    尝试激活窗口: \"" << win.title << L"\"..." << std::endl;
                    if (ForceForegroundWindow(win.hwnd)) {
                        std::wcout << L"    窗口: \"" << win.title << L"\" 已激活。" << std::endl;

                        std::wcout << L"    等待短暂延时 (300ms) 后发送输入..." << std::endl;
                        std::this_thread::sleep_for(std::chrono::milliseconds(300));

                        std::wcout << L"    开始模拟按键和鼠标操作..." << std::endl;

                        std::wcout << L"      发送 'E' 键 (尝试使用扫描码)..." << std::endl;
                        SendKeyWithRandomDelay(0x45, 100, 300, true);
                        std::wcout << L"      按下并释放按键 'E' 完成。" << std::endl;

                        std::this_thread::sleep_for(std::chrono::milliseconds(100));

                        std::wcout << L"      发送鼠标右键点击..." << std::endl;
                        SendRightClickWithRandomDelay(100, 300);
                        std::wcout << L"      模拟鼠标右键点击完成。" << std::endl;
                        std::wcout << L"    窗口 \"" << win.title << L"\" 操作完成。" << std::endl;
                    }
                    else {
                        std::wcerr << L"    !!! 警告: 无法激活窗口: \"" << win.title << L"\". 跳过当前窗口操作。" << std::endl;
                    }
                    std::wcout << L"  ------------------------------------" << std::endl;
                    std::this_thread::sleep_for(std::chrono::milliseconds(500));
                }
            }
            std::wcout << L"--- 桌面 " << (i + 1) << L" 处理完毕 ---" << std::endl;
            if (desktopCount == 1 && i == 0) break;
        }

        std::wcout << L"\n已完成对所有虚拟桌面上所有窗口的一轮操作。" << std::endl;
        double roundDelaySeconds = delayDist(gen_desktop);
        std::wcout << L"下一轮完整扫描开始前等待: " << roundDelaySeconds << L" 秒..." << std::endl;
        std::this_thread::sleep_for(std::chrono::duration<double>(roundDelaySeconds));
        std::wcout << L"\n===========================================================\n" << std::endl;
    }

    ShutdownVirtualDesktopManager();
    return 0;
}