#include "pch.h"
#include "LoginDialog.h"

DUI_BEGIN_MESSAGE_MAP(LoginDialog, WindowImplBase)
DUI_ON_MSGTYPE_CTRNAME(DUI_MSGTYPE_CLICK, _T("closeBtn"), onCloseBtnClick)
DUI_ON_MSGTYPE_CTRNAME(DUI_MSGTYPE_CLICK, _T("minBtn"), onMinBtnClick)
DUI_ON_MSGTYPE_CTRNAME(DUI_MSGTYPE_CLICK, _T("loginBtn"), onLoginBtnClick)
DUI_END_MESSAGE_MAP()

LoginDialog::LoginDialog()
{
}

LoginDialog::~LoginDialog()
{
}

DuiLib::CDuiString LoginDialog::GetSkinType()
{
    return _T("DUILIB");
}

CDuiString LoginDialog::GetSkinFile()
{
    DuiLib::CDuiString dsResID;
    dsResID.Format(_T("%d"), IDR_DUILIB_LOGIN);
    return dsResID;
}

LPCTSTR LoginDialog::GetWindowClassName(void) const
{
    return _T("LoginDialog");
}

void LoginDialog::Notify(TNotifyUI & msg)
{
    return WindowImplBase::Notify(msg);
}

void LoginDialog::InitWindow()
{
}

void LoginDialog::onCloseBtnClick(TNotifyUI & msg)
{
    CWindowWnd::SendMessage(WM_CLOSE);
}

void LoginDialog::onMinBtnClick(TNotifyUI & msg)
{
    CWindowWnd::SendMessage(WM_SYSCOMMAND, SC_MINIMIZE, 0);
}

void LoginDialog::onLoginBtnClick(TNotifyUI & msg)
{
    CEditUI *usernameEdit = static_cast<CEditUI *>(m_pm.FindControl(_T("usernameEdit")));
    CEditUI *passwordEdit = static_cast<CEditUI *>(m_pm.FindControl(_T("passwordEdit")));
    if (usernameEdit == nullptr || passwordEdit == nullptr)
        return;
    TCHAR tmp[1024] = { 0 };
    swprintf_s(tmp, _T("您输入的用户名是 %s, 密码是 %s"), usernameEdit->GetText().GetData(), passwordEdit->GetText().GetData());
    ::MessageBox(m_hWnd, tmp, _T("提示"), MB_OK);
}
