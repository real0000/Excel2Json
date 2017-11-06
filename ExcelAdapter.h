// ExcelAdapter.h
//
// 2017/10/29 Ian Wu/Real0000
//

#ifndef _EXCEL_ADAPTER_H_
#define _EXCEL_ADAPTER_H_

#include <string>
#include <vector>
#include <map>

#include "xlsxio_read.h"
#include "freexl.h"

class ExcelLoader
{
public:
	ExcelLoader();
	virtual ~ExcelLoader();

	virtual bool load(std::string a_Filepath) = 0;

	bool setSheet(std::string a_SheetName);
	std::string getCell(unsigned int a_Row, unsigned int a_Column);
	
protected:
	virtual bool hasSheet(std::string a_SheetName) = 0;
	virtual void initSheet(std::string a_SheetName) = 0;
	void setCell(unsigned int a_Row, unsigned int a_Column, std::string a_Data);

private:
	typedef std::vector< std::vector<std::string> > SheetTable;
	SheetTable *m_pCurrTable;
	std::map<std::string, SheetTable *> m_SheetMap;
};

class XlsLoader : public ExcelLoader
{
public:
	XlsLoader();
	virtual ~XlsLoader();

	virtual bool load(std::string a_Filepath);
	
protected:
	virtual bool hasSheet(std::string a_SheetName);
	virtual void initSheet(std::string a_SheetName);

private:
	std::map<std::string, unsigned int> m_SheetIndex;
	void *m_pHandle;
};

class XlsxLoader : public ExcelLoader
{
public:
	XlsxLoader();
	virtual ~XlsxLoader();

	virtual bool load(std::string a_Filepath);
	
protected:
	virtual bool hasSheet(std::string a_SheetName);
	virtual void initSheet(std::string a_SheetName);

private:
	std::map<std::string, unsigned int> m_SheetIndex;
	xlsxioreader m_Book;
};

class ExcelAdapter
{
public:
	ExcelAdapter();
	virtual ~ExcelAdapter();

	void clear();

	bool load(std::string a_Filepath);
	bool setSheet(std::string a_SheetName);
	std::string getCell(unsigned int a_Row, unsigned int a_Column);

private:
	ExcelLoader *m_pLoader;
};

#endif