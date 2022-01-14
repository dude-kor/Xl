#include <ole2.h>
#include <iostream>

class CXl {
private:
	HRESULT m_hrCOMInit;						// COM 초기화 여부
	HRESULT m_hrInit = ERROR_INVALID_VARIANT;	// Excel 초기화 여부

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
	///  Excel 초기화 상태에서 Excel을 가시화 합니다.
	/// </summary>
	/// <param name="bVisible">: 가시화 여부</param>
	///<returns>
	///성공적으로 Excel을 가시화 했을 경우, 1을 반환합니다.
	///Variant 초기화가 되지 않았을 경우 -1을 반환합니다.
	///Visible 변환을 실패 했을 경우 -2를 반환합니다.
	///</returns>
	int SetVisible(bool bVisible);

	int Open();

	int AddWorkBooks();

	int AddActiveSheet();

	//XFD1048576 가 최대지만 메모리 리소스 부족으로 가로 세로 3540 정도로 제한합니다.
	//default로 실행 했을 경우 clock() 기준으로 628.000 ms 정도 소요됩니다.
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
	/// 작업한 내용을 저장합니다.
	/// </summary>
	int Save();

	/// <summary>
	/// Excel을 종료합니다.
	/// </summary>
	int Close();

protected:
	/// <summary>
	/// CLSID를 선언하고, VARIANT를 초기화하여 Main Frame을 생성합니다.
	/// </summary>
	/// <returns>
	/// CLSID를 선언하지 못했을 경우, 에러 메시지와 -1을 반환합니다.
	/// <para/>인스턴스를 선언하지 못했을 경우, 에러 메시지와 -2를 반환합니다.
	/// <para/>성공적으로 Main Frame을 생성할 경우, 1을 반환합니다.
	/// </returns>
	int Init();

	// Excel 초기화 상태를 반환합니다.
	HRESULT CheckXlInit() { return m_hrInit; };

public:
	// CXl 클래스 생성자입니다. COM 초기화 상태를 변경합니다.
	CXl();

	// CXl 클래스 파괴자입니다.
	virtual ~CXl();
};