#if !defined(AFX_OPCERRDLG_H__5B282B6E_361E_11D4_80E2_00C04F790F3B__INCLUDED_)
#define AFX_OPCERRDLG_H__5B282B6E_361E_11D4_80E2_00C04F790F3B__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// OPCErrDlg.h : header file
//

/////////////////////////////////////////////////////////////////////////////
// COPCErrDlg dialog

class COPCErrDlg : public CDialog
{
// Construction
public:
	COPCErrDlg(CWnd* pParent = NULL);   // standard constructor

	void UpdateText(char *pMsg);

// Dialog Data
	//{{AFX_DATA(COPCErrDlg)
	enum { IDD = IDD_ERRORDLG };
	CString	m_opcerror;
	//}}AFX_DATA


// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(COPCErrDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(COPCErrDlg)
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_OPCERRDLG_H__5B282B6E_361E_11D4_80E2_00C04F790F3B__INCLUDED_)
