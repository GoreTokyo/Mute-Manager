#include <windows.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <iostream>
#include <fstream>
#include <comdef.h>
#include <chrono>
#include <iomanip>
#include <sstream>

const CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
const IID IID_IAudioEndpointVolume = __uuidof(IAudioEndpointVolume);

std::string GetCurrentDateTime() {
	auto now = std::chrono::system_clock::now();
	auto in_time_t = std::chrono::system_clock::to_time_t(now);
	std::tm tm;
	localtime_s(&tm, &in_time_t);
	std::stringstream ss;
	ss << std::put_time(&tm, "%Y-%m-%d %X");
	return ss.str();
}

void LogMessage(const std::string& message) {
	std::ofstream logFile("MuteManager.log", std::ios_base::app);
	if (logFile.is_open()) {
		logFile << GetCurrentDateTime() << " - " << message << std::endl;
		logFile.close();
	}
}

void PrintAndLogMessage(const std::string& message) {
	std::string dateTimeMessage = GetCurrentDateTime() + " - " + message;
	std::cout << dateTimeMessage << std::endl;
	LogMessage(message);
}

class ComInitializer {
public:
	ComInitializer() {
		HRESULT hr = CoInitialize(nullptr);
		if (FAILED(hr)) {
			throw std::runtime_error("COMライブラリの初期化に失敗しました。");
		}
	}
	~ComInitializer() {
		CoUninitialize();
	}
};

bool GetAudioEndpointVolume(IAudioEndpointVolume** ppVolume) {
	ComInitializer comInit;

	IMMDeviceEnumerator* pEnumerator = nullptr;
	IMMDevice* pDevice = nullptr;
	HRESULT hr = CoCreateInstance(CLSID_MMDeviceEnumerator, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&pEnumerator));
	if (FAILED(hr)) {
		PrintAndLogMessage("デバイス列挙子の作成に失敗しました。");
		return false;
	}

	hr = pEnumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &pDevice);
	if (FAILED(hr)) {
		PrintAndLogMessage("オーディオデバイスの取得に失敗しました。");
		pEnumerator->Release();
		return false;
	}

	hr = pDevice->Activate(IID_IAudioEndpointVolume, CLSCTX_ALL, nullptr, (void**)ppVolume);
	if (FAILED(hr)) {
		PrintAndLogMessage("ボリュームインターフェースの取得に失敗しました。");
		pDevice->Release();
		pEnumerator->Release();
		return false;
	}

	pDevice->Release();
	pEnumerator->Release();
	return true;
}

void CheckMuteStatus() {
	IAudioEndpointVolume* pVolume = nullptr;
	if (!GetAudioEndpointVolume(&pVolume)) return;

	BOOL isMuted = FALSE;
	HRESULT hr = pVolume->GetMute(&isMuted);
	if (FAILED(hr)) {
		PrintAndLogMessage("ミュート状態の取得に失敗しました。");
	}
	else if (isMuted) {
		PrintAndLogMessage("マイクはミュートされています。");
	}
	else {
		PrintAndLogMessage("マイクはミュート解除されています。");
	}

	pVolume->Release();
}

void MuteMicrophone() {
	IAudioEndpointVolume* pVolume = nullptr;
	if (!GetAudioEndpointVolume(&pVolume)) return;

	BOOL isMuted = FALSE;
	HRESULT hr = pVolume->GetMute(&isMuted);
	if (FAILED(hr)) {
		PrintAndLogMessage("ミュート状態の取得に失敗しました。");
	}
	else if (isMuted) {
		PrintAndLogMessage("マイクは既にミュートされています。");
	}
	else {
		hr = pVolume->SetMute(TRUE, nullptr);
		if (FAILED(hr)) {
			PrintAndLogMessage("ミュート設定に失敗しました。");
		}
		else {
			PrintAndLogMessage("マイクをミュートしました。");
		}
	}

	CheckMuteStatus();
	pVolume->Release();
}

void UnmuteMicrophone() {
	IAudioEndpointVolume* pVolume = nullptr;
	if (!GetAudioEndpointVolume(&pVolume)) return;

	BOOL isMuted = FALSE;
	HRESULT hr = pVolume->GetMute(&isMuted);
	if (FAILED(hr)) {
		PrintAndLogMessage("ミュート状態の取得に失敗しました。");
	}
	else if (!isMuted) {
		PrintAndLogMessage("マイクは既にミュート解除されています。");
	}
	else {
		hr = pVolume->SetMute(FALSE, nullptr);
		if (FAILED(hr)) {
			PrintAndLogMessage("ミュート解除に失敗しました。");
		}
		else {
			PrintAndLogMessage("マイクのミュートが解除されました。");
		}
	}

	CheckMuteStatus();
	pVolume->Release();
}

int main() {
	int choice;

	while (true) {
		std::cout << "1: ミュート\n2: ミュート解除\n3: ミュート状態の確認\n4: 終了\n";
		std::cout << "操作を選択してください: ";
		std::cin >> choice;

		try {
			switch (choice) {
			case 1:
				MuteMicrophone();
				break;
			case 2:
				UnmuteMicrophone();
				break;
			case 3:
				CheckMuteStatus();
				break;
			case 4:
				PrintAndLogMessage("プログラムを終了します。");
				return 0;
			default:
				PrintAndLogMessage("無効な選択がされました。");
			}
		}
		catch (const std::exception& e) {
			PrintAndLogMessage(e.what());
		}
	}

	return 0;
}

