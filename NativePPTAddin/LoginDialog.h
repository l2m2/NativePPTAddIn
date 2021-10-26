#pragma once

class LoginDialog : public WindowImplBase
{
public:
    LoginDialog();
    ~LoginDialog();

public:
    DuiLib::CDuiString GetSkinType() override;
    CDuiString GetSkinFile() override;
    LPCTSTR GetWindowClassName(void) const override;
    void Notify(TNotifyUI &msg) override;
    void InitWindow() override;

protected:
    DUI_DECLARE_MESSAGE_MAP()
    void onCloseBtnClick(TNotifyUI& msg);
    void onMinBtnClick(TNotifyUI& msg);
    void onLoginBtnClick(TNotifyUI& msg);
};
