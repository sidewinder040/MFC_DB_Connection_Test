
// MFC_DBDlg.cpp : implementation file
//

#include "pch.h"
#include "framework.h"
#include "MFC_DB.h"
#include "MFC_DBDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CMFCDBDlg dialog



CMFCDBDlg::CMFCDBDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_MFC_DB_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMFCDBDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_ListControl);
	DDX_Control(pDX, IDC_BUTTON1, m_buttonRead);
}

BEGIN_MESSAGE_MAP(CMFCDBDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CMFCDBDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CMFCDBDlg message handlers

BOOL CMFCDBDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here

	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CMFCDBDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CMFCDBDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CMFCDBDlg::OnBnClickedButton1()
{
	// TODO: Add your control notification handler code here
	CDatabase database;
	CString SqlString;
	CString strID, strName, strAge;
	/*CString sDriver = L"ODBC Driver 18 for SQL Server";*/
	//CString sFile = L"D:\\Test.mdb";
	CString sDsn = L"myUserDSN";
	CString sDatabase = L"Test";
	// You must change above path if it's different
	int iRec = 0;

	// Build ODBC connection string
	//sDsn.Format(L"ODBC;DRIVER={%s};DSN='%s';database=%s", sDriver, sDsn, sDatabase);
	sDsn = L"DSN=myUserDSN; Trusted_Connection=Yes;";

	TRY{
		// Open the database
		database.Open(NULL,false,false,sDsn);

	// Allocate the recordset
	CRecordset recset(&database);
	// Build the SQL statement
	SqlString = "SELECT ID, Name, Age " "FROM Employees";

	// Execute the query

	recset.Open(CRecordset::forwardOnly, SqlString, CRecordset::readOnly);
	// Reset List control if there is any data
	ResetListControl();
	// populate Grids
	ListView_SetExtendedListViewStyle(m_ListControl, LVS_EX_GRIDLINES);

	// Column width and heading
	m_ListControl.InsertColumn(0, L"Emp ID", LVCFMT_LEFT, -1, 0);
	m_ListControl.InsertColumn(1, L"Name", LVCFMT_LEFT, -1, 1);
	m_ListControl.InsertColumn(2, L"Age", LVCFMT_LEFT, -1, 1);
	m_ListControl.SetColumnWidth(0, 120);
	m_ListControl.SetColumnWidth(1, 200);
	m_ListControl.SetColumnWidth(2, 200);

	// Loop through each record
	while (!recset.IsEOF()) {
		// Copy each column into a variable
		recset.GetFieldValue(L"ID", strID);
		recset.GetFieldValue(L"Name", strName);
		recset.GetFieldValue(L"Age", strAge);

		// Insert values into the list control
		iRec = m_ListControl.InsertItem(0, strID, 0);
		m_ListControl.SetItemText(0, 1, strName);
		m_ListControl.SetItemText(0, 2, strAge);

		// goto next record
		recset.MoveNext();
	}
	// Close the database
	database.Close();
}CATCH(CDBException, e) {
	// If a database exception occured, show error msg
	AfxMessageBox(L"Database error: " + e->m_strError);
}
END_CATCH;
}

// Reset List control
void CMFCDBDlg::ResetListControl() {
	m_ListControl.DeleteAllItems();
	int iNbrOfColumns;
	CHeaderCtrl* pHeader = (CHeaderCtrl*)m_ListControl.GetDlgItem(0);
	if (pHeader) {
		iNbrOfColumns = pHeader->GetItemCount();
	}
	for (int i = iNbrOfColumns; i >= 0; i--) {
		m_ListControl.DeleteColumn(i);
	}
}
	
