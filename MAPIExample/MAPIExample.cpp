// MAPIExample.cpp : コンソール アプリケーションのエントリ ポイントを定義します。
//

#include "stdafx.h"

class MapiMail
{
public:
    MapiMail()
        : mapi_dll_(LoadLibraryW(L"mapi32.dll"), &FreeLibrary)
        , MAPILogon_(nullptr)
        , MAPILogoff_(nullptr)
        , MAPISendMailW_(nullptr)
        , session_(0)
    {
        if (!mapi_dll_)
            throw std::runtime_error("cannot load library");

        MAPILogon_ = reinterpret_cast<LPMAPILOGON>(GetProcAddress(mapi_dll_.get(), "MAPILogon"));
        MAPILogoff_ = reinterpret_cast<LPMAPILOGOFF>(GetProcAddress(mapi_dll_.get(), "MAPILogoff"));
        MAPISendMailW_ = reinterpret_cast<LPMAPISENDMAILW>(GetProcAddress(mapi_dll_.get(), "MAPISendMailW"));
        if (!MAPILogon_ || !MAPILogoff_ || !MAPISendMailW_)
            throw std::runtime_error("cannot get function pointer");

        auto error = MAPILogon_(NULL, NULL, NULL, MAPI_NEW_SESSION, 0, &session_);
        if (FAILED(error))
            throw std::runtime_error("error code: " + std::to_string(error));

    }

    ~MapiMail()
    {
        if (session_ == 0)
            return;

        auto error = MAPILogoff_(session_, 0, 0, 0);
        if (FAILED(error))
            throw std::runtime_error("error code: " + std::to_string(error));
    }

    void Send(MapiMessageW& message)
    {
        if (session_ == 0)
            throw std::runtime_error("failed to initialization");

        auto error = MAPISendMailW_(session_, 0, &message, MAPI_DIALOG, 0);
        if (FAILED(error))
            throw std::runtime_error("error code: " + std::to_string(error));
    }

private:
    std::unique_ptr<std::remove_pointer<HMODULE>::type, decltype(&FreeLibrary)> mapi_dll_;
    LPMAPILOGON MAPILogon_;
    LPMAPILOGOFF MAPILogoff_;
    LPMAPISENDMAILW MAPISendMailW_;
    LHANDLE session_;
};

void MAPIExample()
{
    std::vector<MapiRecipDescW> recip_desc {
        { 0, MAPI_TO, L"いべ", L"t.ibe@adacotech.co.jp" },
    };
    std::vector<MapiFileDescW> file_desc {
        { 0, 0, static_cast<ULONG>(-1), LR"(ReadMe.txt)", LR"(お読みください.txt)", nullptr },
    };
    MapiMessageW message {
        0, L"Hello", L"Hi!\nこんにちは。",
        nullptr, nullptr, nullptr, 0, nullptr,
        recip_desc.size(), recip_desc.data(), file_desc.size(), file_desc.data()
    };

    MapiMail mail;
    mail.Send(message);
}

int main()
{
    try {
        MAPIExample();
    }
    catch (std::exception& e) {
        std::cerr << e.what() << std::endl;
        return 1;
    }

    return 0;
}
