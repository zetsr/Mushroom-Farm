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
            std::wcerr << L"����: �޷���������� 0x" << std::hex << vk << std::dec
                << L" ӳ�䵽ɨ���롣�����˵���ʹ��������롣" << std::endl;
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
        std::wcerr << L"SendInput (����) ʧ�ܡ��������: " << GetLastError() << std::endl;
    }
    std::uniform_int_distribution<> dis(minMs, maxMs);
    int delay = dis(gen_desktop);
    std::this_thread::sleep_for(std::chrono::milliseconds(delay));
    if (SendInput(1, &inputUp, sizeof(INPUT)) == 0) {
        std::wcerr << L"SendInput (�ͷ�) ʧ�ܡ��������: " << GetLastError() << std::endl;
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
            std::wcerr << L"����: AttachThreadInput(TRUE) ʧ�ܡ�����: " << GetLastError() << std::endl;
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
            std::wcerr << L"����: AttachThreadInput(FALSE) ʧ�ܡ�����: " << GetLastError() << std::endl;
        }
    }
    std::this_thread::sleep_for(std::chrono::milliseconds(250));
    return (GetForegroundWindow() == hwnd);
}

IVirtualDesktopManagerInternal* g_pVDManagerInternal = nullptr;

bool InitializeVirtualDesktopManager() {
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        std::wcerr << L"CoInitializeEx ʧ�ܡ�HR: 0x" << std::hex << hr << std::endl;
        return false;
    }

    IServiceProvider* pServiceProvider = nullptr;
    hr = CoCreateInstance(CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER, IID_IServiceProvider, (void**)&pServiceProvider);
    if (FAILED(hr)) {
        std::wcerr << L"CoCreateInstance(CLSID_ImmersiveShell) ʧ�ܡ�HR: 0x" << std::hex << hr << std::endl;
        CoUninitialize();
        return false;
    }

    hr = pServiceProvider->QueryService(CLSID_VirtualDesktopManager_Service, IID_IVirtualDesktopManagerInternal, (void**)&g_pVDManagerInternal);
    if (FAILED(hr)) {
        const GUID IID_IVirtualDesktopManagerInternal_Old = { 0xF31574D6, 0xB682, 0x4CDC, {0xBD, 0x56, 0x18, 0x27, 0x86, 0x0A, 0xBE, 0xC6} };
        hr = pServiceProvider->QueryService(CLSID_VirtualDesktopManager_Service, IID_IVirtualDesktopManagerInternal_Old, (void**)&g_pVDManagerInternal);
        if (FAILED(hr)) {
            std::wcerr << L"QueryService(CLSID_VirtualDesktopManager_Service) ʧ�� (�¾�IID������ʧ��)��HR: 0x" << std::hex << hr << std::endl;
            SafeRelease(&pServiceProvider);
            CoUninitialize();
            return false;
        }
        std::wcout << L"ʹ���˾ɰ� IVirtualDesktopManagerInternal IID��" << std::endl;
    }

    SafeRelease(&pServiceProvider);
    if (!g_pVDManagerInternal) {
        CoUninitialize();
        return false;
    }
    std::wcout << L"���������������ʼ���ɹ���" << std::endl;
    return true;
}

void ShutdownVirtualDesktopManager() {
    SafeRelease(&g_pVDManagerInternal);
    CoUninitialize();
    std::wcout << L"��������������ѹرա�" << std::endl;
}

bool SwitchToDesktopByIndex(UINT index) {
    if (!g_pVDManagerInternal) return false;

    IObjectArray* pDesktopsArray = nullptr;
    HRESULT hr = g_pVDManagerInternal->GetDesktops(&pDesktopsArray);
    if (FAILED(hr) || !pDesktopsArray) {
        std::wcerr << L"GetDesktops ʧ�ܡ�HR: 0x" << std::hex << hr << std::endl;
        return false;
    }

    IVirtualDesktop* pDesktopToSwitch = nullptr;
    hr = pDesktopsArray->GetAt(index, __uuidof(IVirtualDesktop), (void**)&pDesktopToSwitch);
    if (FAILED(hr) || !pDesktopToSwitch) {
        std::wcerr << L"GetAt(" << index << L") ʧ�ܡ�HR: 0x" << std::hex << hr << std::endl;
        SafeRelease(&pDesktopsArray);
        return false;
    }

    hr = g_pVDManagerInternal->SwitchDesktop(pDesktopToSwitch);
    if (FAILED(hr)) {
        std::wcerr << L"SwitchDesktop(" << index << L") ʧ�ܡ�HR: 0x" << std::hex << hr << std::endl;
    }

    SafeRelease(&pDesktopToSwitch);
    SafeRelease(&pDesktopsArray);
    return SUCCEEDED(hr);
}

int main() {
    SetConsoleOutputCP(CP_UTF8);
    std::locale::global(std::locale(""));

    if (!InitializeVirtualDesktopManager()) {
        std::wcerr << L"�޷���ʼ��������������������򽫽��ڵ�ǰ�������У�������ã���" << std::endl;
    }

    std::uniform_real_distribution<> delayDist(5.0, 10.0);
    std::wcout << L"������������ʼ������ѭ������..." << std::endl;

    while (true) {
        UINT desktopCount = 0;
        if (g_pVDManagerInternal) {
            g_pVDManagerInternal->GetCount(&desktopCount);
        }
        else {
            desktopCount = 1;
            std::wcout << L"�������������δ��ʼ������ɨ�赱ǰ���档" << std::endl;
        }

        if (desktopCount == 0 && g_pVDManagerInternal) {
            std::wcout << L"δ�ҵ��������棨����GetCount����0�����ȴ�������..." << std::endl;
            std::this_thread::sleep_for(std::chrono::seconds(5));
            continue;
        }

        std::wcout << L"��ʼ��һ��ɨ�裬������ " << desktopCount << L" ���������档" << std::endl;

        for (UINT i = 0; i < desktopCount; ++i) {
            if (g_pVDManagerInternal) {
                std::wcout << L"\n--- �����л������� " << (i + 1) << L" ---" << std::endl;
                if (!SwitchToDesktopByIndex(i)) {
                    std::wcerr << L"�л������� " << (i + 1) << L" ʧ�ܣ����������档" << std::endl;
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
                std::wcout << L"���� " << (i + 1) << L" ��δ�ҵ��κ� Roblox ���ڡ�" << std::endl;
            }
            else {
                std::wcout << L"������ " << (i + 1) << L" ���ҵ� " << windowsOnThisDesktop.size() << L" �� Roblox ���ڣ���ʼ����..." << std::endl;
                for (const auto& win : windowsOnThisDesktop) {
                    std::wcout << L"  ׼����������: \"" << win.title << L"\" (λ������ " << (i + 1) << L")" << std::endl;

                    std::wcout << L"    ���Լ����: \"" << win.title << L"\"..." << std::endl;
                    if (ForceForegroundWindow(win.hwnd)) {
                        std::wcout << L"    ����: \"" << win.title << L"\" �Ѽ��" << std::endl;

                        std::wcout << L"    �ȴ�������ʱ (300ms) ��������..." << std::endl;
                        std::this_thread::sleep_for(std::chrono::milliseconds(300));

                        std::wcout << L"    ��ʼģ�ⰴ����������..." << std::endl;

                        std::wcout << L"      ���� 'E' �� (����ʹ��ɨ����)..." << std::endl;
                        SendKeyWithRandomDelay(0x45, 100, 300, true);
                        std::wcout << L"      ���²��ͷŰ��� 'E' ��ɡ�" << std::endl;

                        std::this_thread::sleep_for(std::chrono::milliseconds(100));

                        std::wcout << L"      ��������Ҽ����..." << std::endl;
                        SendRightClickWithRandomDelay(100, 300);
                        std::wcout << L"      ģ������Ҽ������ɡ�" << std::endl;
                        std::wcout << L"    ���� \"" << win.title << L"\" ������ɡ�" << std::endl;
                    }
                    else {
                        std::wcerr << L"    !!! ����: �޷������: \"" << win.title << L"\". ������ǰ���ڲ�����" << std::endl;
                    }
                    std::wcout << L"  ------------------------------------" << std::endl;
                    std::this_thread::sleep_for(std::chrono::milliseconds(500));
                }
            }
            std::wcout << L"--- ���� " << (i + 1) << L" ������� ---" << std::endl;
            if (desktopCount == 1 && i == 0) break;
        }

        std::wcout << L"\n����ɶ������������������д��ڵ�һ�ֲ�����" << std::endl;
        double roundDelaySeconds = delayDist(gen_desktop);
        std::wcout << L"��һ������ɨ�迪ʼǰ�ȴ�: " << roundDelaySeconds << L" ��..." << std::endl;
        std::this_thread::sleep_for(std::chrono::duration<double>(roundDelaySeconds));
        std::wcout << L"\n===========================================================\n" << std::endl;
    }

    ShutdownVirtualDesktopManager();
    return 0;
}