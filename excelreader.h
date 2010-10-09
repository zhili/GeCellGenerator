#pragma once

// #include <standard library headers>
#include <string>
#include <list>

// #include <other library headers>
#include <ole2.h>
// #include <other library headers>
#include <comutil.h>
#include <Shlwapi.h>
#include <atlsafe.h>


template<typename T> inline std::wstring towstr(const T &x){
	std::wostringstream o;
	if (!(o << x)){
		// Log Error Message here
	}
	return o.str();
}



template<typename T> inline std::string tostr(const T &x){
	std::ostringstream o;
	if (!(o << x)){
		// Log Error Message here
	}
	return o.str();
}

template<> inline std::wstring towstr<std::string>(const std::string &x){
	std::wstring wstrReturn;
 
	wstrReturn.assign(x.begin(), x.end());
 
	return wstrReturn;
}
 

template<> inline std::string tostr<std::wstring>(const std::wstring &x){
	std::string strReturn;
 
	strReturn.assign(x.begin(), x.end());
 
	return strReturn;
}

template<> inline std::wstring towstr<variant_t>(const variant_t &x){
	return std::wstring((_bstr_t)(_variant_t)x);
}

template<> inline std::string tostr<variant_t>(const variant_t &x){
	return std::string((_bstr_t)(_variant_t)x);
}


namespace msexcel {

	
	class excelreader {
	 /**
     * description
     */
    class ExcelException : public std::exception
    {
    public:
        ExcelException(const std::string& s = "") : msg(s) {}
        virtual ~ExcelException() throw() {}

    public:
        virtual const char* what() const throw() { return msg.c_str(); }
        virtual const std::string What() const throw() { return msg; }

    private:
        std::string msg;

    };

	public:
		excelreader(const std::wstring& filename = L"", const std::wstring& sheet = L"", bool visible = false);
		virtual ~excelreader();

	public:
		int rowCount();
		int colCount();
		void SetData(const std::wstring& cell,
						 const std::wstring& data);     // set a range
		void SetData(int nRow, int nBeg,
						 const std::list<std::wstring>& data); // set a line
		void SetData(const std::wstring& colTag, int nBeg,
						 const std::list<std::wstring>& data); // set a column
		void SaveAsFile(const std::wstring& filename);
		double excelreader::GetDataAsDouble(int row, int col);
		std::string excelreader::GetDataAsString(int row, int col);
	private:
		void CreateNewWorkbook();
		void OpenWorkbook(const std::wstring& filename, const std::wstring& sheet);
		void GetActiveSheet();
		void GetSheet(const std::wstring& sheetname);
		std::wstring GetColumnName(int nIndex);
		bool CheckFilename(const std::wstring& filename);
		void InstanceExcelWorkbook(int visible);

		static void excelreader::AutoWrap(int autoType, VARIANT *pvResult,
										IDispatch *pDisp, LPOLESTR ptName,
										int cArgs...);

	private:
		IDispatch* pXlApp;
		IDispatch* pXlBooks;
		IDispatch* pXlBook;
		IDispatch* pXlSheet;

	};
}