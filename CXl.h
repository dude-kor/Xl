<<<<<<< HEAD
#include <ole2.h>
#include <iostream>

class CXl {
private:
	HRESULT m_hrCOMInit;						// COM �ʱ�ȭ ����
	HRESULT m_hrInit = ERROR_INVALID_VARIANT;	// Excel �ʱ�ȭ ����

	struct PROPERTIES {
		IDispatch* pXlApp;		// Excel
		IDispatch* pXlBooks;	// Main Frame
		IDispatch* pXlBook;		// Work Book
		IDispatch* pXlSheet;	// Sheet
		IDispatch* pXlRange;	// Range

		VARIANT pvaArray;
	} m_XlProp;

public:
	/// <summary>
	///  Excel �ʱ�ȭ ���¿��� Excel�� ����ȭ �մϴ�.
	/// </summary>
	/// <param name="bVisible">: ����ȭ ����</param>
	///<returns>
	///���������� Excel�� ����ȭ ���� ���, 1�� ��ȯ�մϴ�.
	///Variant �ʱ�ȭ�� ���� �ʾ��� ��� -1�� ��ȯ�մϴ�.
	///Visible ��ȯ�� ���� ���� ��� -2�� ��ȯ�մϴ�.
	///</returns>
	int SetVisible(bool bVisible);

	int Open();

	int AddWorkBooks();

	int AddActiveSheet();

	//XFD1048576 �� �ִ����� �޸� ���ҽ� �������� ���� ���� 3540 ������ �����մϴ�.
	//default�� ���� ���� ��� clock() �������� 628.000 ms ���� �ҿ�˴ϴ�.
	int SetRange(const char* pcszStart = "A1", const char* pcszEnd = "EFD3540");

	int SetSafeBound(unsigned int unRowMin = 1,
		unsigned int unRowMax = 3540,
		unsigned int unColumnMin = 1,
		unsigned int unColumnMax = 3540);
	int SetSafeBound(const char* pcszStart = "A1", const char* pcszEnd = "EFD3540");

	int SetData(const char* pcszData, long lRow, long lColumn);

	int Print();

	int SetActiveSheet(int nSheet);

	/// <summary>
	/// �۾��� ������ �����մϴ�.
	/// </summary>
	int Save();

	/// <summary>
	/// Excel�� �����մϴ�.
	/// </summary>
	int Close();

protected:
	/// <summary>
	/// CLSID�� �����ϰ�, VARIANT�� �ʱ�ȭ�Ͽ� Main Frame�� �����մϴ�.
	/// </summary>
	/// <returns>
	/// CLSID�� �������� ������ ���, ���� �޽����� -1�� ��ȯ�մϴ�.
	/// <para/>�ν��Ͻ��� �������� ������ ���, ���� �޽����� -2�� ��ȯ�մϴ�.
	/// <para/>���������� Main Frame�� ������ ���, 1�� ��ȯ�մϴ�.
	/// </returns>
	int Init();

	// Excel �ʱ�ȭ ���¸� ��ȯ�մϴ�.
	HRESULT CheckXlInit() { return m_hrInit; };

public:
	// CXl Ŭ���� �������Դϴ�. COM �ʱ�ȭ ���¸� �����մϴ�.
	CXl();

	// CXl Ŭ���� �ı����Դϴ�.
	virtual ~CXl();
=======
#include <ole2.h>
#include <iostream>

class CXl {
private:
	HRESULT m_hrCOMInit;						// COM �ʱ�ȭ ����
	HRESULT m_hrInit = ERROR_INVALID_VARIANT;	// Excel �ʱ�ȭ ����

	struct PROPERTIES {
		IDispatch* pXlApp;		// Excel
		IDispatch* pXlBooks;	// Main Frame
		IDispatch* pXlBook;		// Work Book
		IDispatch* pXlSheet;	// Sheet
		IDispatch* pXlRange;	// Range

		VARIANT pvaArray;
	} m_XlProp;

public:
	/// <summary>
	///  Excel �ʱ�ȭ ���¿��� Excel�� ����ȭ �մϴ�.
	/// </summary>
	/// <param name="bVisible">: ����ȭ ����</param>
	///<returns>
	///���������� Excel�� ����ȭ ���� ���, 1�� ��ȯ�մϴ�.
	///Variant �ʱ�ȭ�� ���� �ʾ��� ��� -1�� ��ȯ�մϴ�.
	///Visible ��ȯ�� ���� ���� ��� -2�� ��ȯ�մϴ�.
	///</returns>
	int SetVisible(bool bVisible);

	int Open();

	int AddWorkBooks();

	int AddActiveSheet();

	//XFD1048576 �� �ִ����� �޸� ���ҽ� �������� ���� ���� 3540 ������ �����մϴ�.
	//default�� ���� ���� ��� clock() �������� 628.000 ms ���� �ҿ�˴ϴ�.
	int SetRange(const char* pcszStart = "A1", const char* pcszEnd = "EFD3540");

	int SetSafeBound(unsigned int unRowMin = 1,
		unsigned int unRowMax = 3540,
		unsigned int unColumnMin = 1,
		unsigned int unColumnMax = 3540);
	int SetSafeBound(const char* pcszStart = "A1", const char* pcszEnd = "EFD3540");

	int SetData(const char* pcszData, long lRow, long lColumn);

	int Print();

	int SetActiveSheet(int nSheet);

	/// <summary>
	/// �۾��� ������ �����մϴ�.
	/// </summary>
	int Save();

	/// <summary>
	/// Excel�� �����մϴ�.
	/// </summary>
	int Close();

protected:
	/// <summary>
	/// CLSID�� �����ϰ�, VARIANT�� �ʱ�ȭ�Ͽ� Main Frame�� �����մϴ�.
	/// </summary>
	/// <returns>
	/// CLSID�� �������� ������ ���, ���� �޽����� -1�� ��ȯ�մϴ�.
	/// <para/>�ν��Ͻ��� �������� ������ ���, ���� �޽����� -2�� ��ȯ�մϴ�.
	/// <para/>���������� Main Frame�� ������ ���, 1�� ��ȯ�մϴ�.
	/// </returns>
	int Init();

	// Excel �ʱ�ȭ ���¸� ��ȯ�մϴ�.
	HRESULT CheckXlInit() { return m_hrInit; };

public:
	// CXl Ŭ���� �������Դϴ�. COM �ʱ�ȭ ���¸� �����մϴ�.
	CXl();

	// CXl Ŭ���� �ı����Դϴ�.
	virtual ~CXl();
>>>>>>> 2b52f65a308d6593fa542001cff22d180e17066f
};