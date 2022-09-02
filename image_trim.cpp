#include <iostream>
#include <opencv2/opencv.hpp>
#include <Windows.h>
#include <ShObjIdl.h>
#include <atlconv.h>
#include <vector>

#ifdef _DEBUG
#pragma comment(lib, "opencv_world460d.lib")
#else
#pragma comment(lib, "opencv_world460.lib")
#endif

// std::wstringをstd::stringへ変換
// cf. https://www.yasuhisay.info/entry/20090722/1248245439
void narrow(const std::wstring& src, std::string& dest) {
    size_t i;
    size_t mbsSz = src.length() * MB_CUR_MAX + 1;
    char* mbs = new char[mbsSz];
    wcstombs_s(&i, mbs, mbsSz, src.c_str(), src.length() * MB_CUR_MAX + 1);
    dest = mbs;
    delete[] mbs;
}

// フォルダのパスを取得するダイアログを表示
std::string open_folder_dialog()
{
    PWSTR pszFolderPath = NULL;

    IFileOpenDialog* pFileOpen;
    DWORD options;
    HRESULT hr;

    hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL, IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));
    pFileOpen->GetOptions(&options);
    pFileOpen->SetOptions(options | FOS_PICKFOLDERS);
    if (SUCCEEDED(hr))
    {
        hr = pFileOpen->Show(NULL);
        if (SUCCEEDED(hr))
        {
            IShellItem* pItem;

            hr = pFileOpen->GetResult(&pItem);
            if (SUCCEEDED(hr))
            {
                hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFolderPath);
                if (SUCCEEDED(hr))
                {
                    CoTaskMemFree(pszFolderPath);
                }
                pItem->Release();
            }
        }
        pFileOpen->Release();
    }

    std::string res;
    narrow(std::wstring(pszFolderPath), res);
    return res;
}

// 複数のファイルパスを取得するダイアログを表示
std::vector<std::string> open_files_dialog()
{
    std::vector<std::string> vFilePath;
    HRESULT hr;
    IFileOpenDialog* pFileOpen;
    DWORD options;

    hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL, IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));
    if (SUCCEEDED(hr))
    {
        pFileOpen->GetOptions(&options);
        pFileOpen->SetOptions(options | FOS_FORCEFILESYSTEM | FOS_ALLOWMULTISELECT);
        hr = pFileOpen->Show(NULL);
        if (SUCCEEDED(hr))
        {
            IShellItemArray* pItemArray;
            hr = pFileOpen->GetResults(&pItemArray);
            if (SUCCEEDED(hr))
            {
                IEnumShellItems* pEnum;
                IShellItem* pItem;
                ULONG fetched;
                hr = pItemArray->EnumItems(&pEnum);
                do {
                    PWSTR pszFilePath;
                    hr = pEnum->Next(1, &pItem, &fetched);
                    hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);
                    if (SUCCEEDED(hr) && fetched > 0)
                    {
                        std::string res;
                        narrow(std::wstring(pszFilePath), res);
                        vFilePath.push_back(res);
                    }
                } while (fetched > 0);
                pItem->Release();
            }
        }
        pFileOpen->Release();
    }

    return vFilePath;
}

int main()
{
    // COM初期化
    if (FAILED(CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE))) return -1;

    // ロケール設定
    setlocale(LC_ALL, "Japanese");

    // 操作元ファイルを指定
    std::vector<std::string> fileNames = open_files_dialog();

    // 保存先フォルダを指定
    std::string folderName = open_folder_dialog();

    // ソースファイルから画像を切り取り、目的フォルダに保存
    for (const auto& file : fileNames) {
        cv::Mat src_img, dst_img;
        src_img = cv::imread(file);
        dst_img = src_img(cv::Range(100, 1050), cv::Range(1920, 3840));
        cv::imwrite(folderName + "\\" + file.substr(file.find_last_of("\\") + 1), dst_img);
    }

    // COM後処理
    CoUninitialize();

    return 0;
}
