// ExcelAdapter.cpp
//
// 2017/10/29 Ian Wu/Real0000
//
#include <string>
#include <windows.h>
#include <assert.h>

#include "ExcelAdapter.h"

#pragma region ExcelLoader
//
// ExcelLoader
//
ExcelLoader::ExcelLoader()
	: m_pCurrTable(nullptr)
{
}

ExcelLoader::~ExcelLoader()
{
	for( auto it = m_SheetMap.begin() ; it != m_SheetMap.end() ; ++it ) delete it->second;
	m_SheetMap.clear();
}

bool ExcelLoader::setSheet(std::string a_SheetName)
{
	auto it = m_SheetMap.find(a_SheetName);
	if( m_SheetMap.end() != it )
	{
		m_pCurrTable = it->second;
		return true;
	}

	if( hasSheet(a_SheetName) )
	{
		m_pCurrTable = new SheetTable();
		m_SheetMap.insert(std::make_pair(a_SheetName, m_pCurrTable));
		initSheet(a_SheetName);
		return true;
	}

	return false;
}

std::string ExcelLoader::getCell(unsigned int a_Row, unsigned int a_Column)
{
	if( nullptr == m_pCurrTable ||
		a_Row >= m_pCurrTable->size() ||
		a_Column >= m_pCurrTable->at(a_Row).size() ) return "";
	return (*m_pCurrTable)[a_Row][a_Column];
}
	
void ExcelLoader::setCell(unsigned int a_Row, unsigned int a_Column, std::string a_Data)
{
	assert(nullptr != m_pCurrTable);
	while( a_Row >= m_pCurrTable->size() ) m_pCurrTable->push_back(std::vector<std::string>());
	if( a_Column >= m_pCurrTable->at(a_Row).size() ) m_pCurrTable->at(a_Row).resize(a_Column + 1, "");
	(*m_pCurrTable)[a_Row][a_Column] = a_Data;
}
#pragma endregion

#pragma region XlsLoader
//
// XlsLoader
//
XlsLoader::XlsLoader()
	: m_pHandle(nullptr)
{
}

XlsLoader::~XlsLoader()
{
	if( nullptr != m_pHandle ) freexl_close(m_pHandle);
}

bool XlsLoader::load(std::string a_Filepath)
{
	if( FREEXL_OK != freexl_open(a_Filepath.c_str(), (const void **)&m_pHandle) ) return false;

	unsigned int l_NumSheet = 0;
	if( FREEXL_OK != freexl_get_info(m_pHandle, FREEXL_BIFF_SHEET_COUNT, &l_NumSheet) ) return false;

	for( unsigned int i=0 ; i<l_NumSheet ; ++i )
	{
		const char *l_pName = nullptr;
		if( FREEXL_OK != freexl_get_worksheet_name(m_pHandle, i, &l_pName) ) continue;
		m_SheetIndex.insert(std::make_pair(std::string(l_pName), i));
	}
}
	
bool XlsLoader::hasSheet(std::string a_SheetName)
{
	return m_SheetIndex.end() != m_SheetIndex.find(a_SheetName);
}

void XlsLoader::initSheet(std::string a_SheetName)
{
	freexl_select_active_worksheet(m_pHandle, m_SheetIndex[a_SheetName]);

	unsigned int l_Row = 0;
	unsigned short l_Column = 0;
	freexl_worksheet_dimensions(m_pHandle, &l_Row, &l_Column);
	if( 0 == l_Row || 0 == l_Column ) return;

	char l_Buff[32];
	for( int r=l_Row-1 ; r>=0 ; --r )
	{
		for( int c=l_Column-1 ; c>=0 ; --c )
		{
			FreeXL_CellValue l_Val;
			freexl_get_cell_value(m_pHandle, r, c, &l_Val);
			switch( l_Val.type )
			{
				case FREEXL_CELL_INT:
					snprintf(l_Buff, 32, "%d", l_Val.value.int_value);
					setCell(r, c, l_Buff);
					break;

				case FREEXL_CELL_DOUBLE:
					snprintf(l_Buff, 32, "%f", l_Val.value.double_value);
					setCell(r, c, l_Buff);
					break;

				case FREEXL_CELL_TEXT:
				case FREEXL_CELL_SST_TEXT:
				case FREEXL_CELL_DATE:
				case FREEXL_CELL_DATETIME:
				case FREEXL_CELL_TIME:
					setCell(r, c, l_Val.value.text_value);
					break;

				//case FREEXL_CELL_NULL:
				default:break;
			}
		}
	}
}
#pragma endregion

#pragma region XlsxLoader
//
// XlsxLoader
//
XlsxLoader::XlsxLoader()
	: m_Book(nullptr)
{
}

XlsxLoader::~XlsxLoader()
{
	if( nullptr != m_Book ) xlsxioread_close(m_Book);
}

bool XlsxLoader::load(std::string a_Filepath)
{
	m_Book = xlsxioread_open(a_Filepath.c_str());
	if( nullptr == m_Book ) return false;

	xlsxioreadersheetlist l_SheetList = xlsxioread_sheetlist_open(m_Book);
	if( nullptr == l_SheetList ) return false;

	const char *l_pSheetName = nullptr;
	unsigned int l_Index = 0;
	while( nullptr != (l_pSheetName = xlsxioread_sheetlist_next(l_SheetList)) )
	{
		m_SheetIndex[l_pSheetName] = l_Index;
		++l_Index;
	}
	xlsxioread_sheetlist_close(l_SheetList);

	return true;
}
	
bool XlsxLoader::hasSheet(std::string a_SheetName)
{
	return m_SheetIndex.end() != m_SheetIndex.find(a_SheetName);
}

void XlsxLoader::initSheet(std::string a_SheetName)
{
	xlsxioreadersheet l_Sheet = xlsxioread_sheet_open(m_Book, a_SheetName.c_str(), XLSXIOREAD_SKIP_NONE);
	unsigned int l_Row = 0, l_Column = 0;
	while( 0 != xlsxioread_sheet_next_row(l_Sheet) )
	{
		char *l_pBuff = nullptr;
		while( nullptr != (l_pBuff = xlsxioread_sheet_next_cell(l_Sheet)) )
		{
			setCell(l_Row, l_Column, l_pBuff);
			++l_Column;
			free(l_pBuff);
		}
		++l_Row;
		l_Column = 0;
	}
	xlsxioread_sheet_close(l_Sheet);
}
#pragma endregion

#pragma region ExcelAdapter
//
// ExcelAdapter
//
ExcelAdapter::ExcelAdapter()
	: m_pLoader(nullptr)
{
}

ExcelAdapter::~ExcelAdapter()
{
	clear();
}

void ExcelAdapter::clear()
{
	if( nullptr != m_pLoader ) delete m_pLoader;
	m_pLoader = nullptr;
}

bool ExcelAdapter::load(std::string a_Filepath)
{
	clear();
	std::string l_Ext("");
	{
		size_t l_Idx = a_Filepath.find_last_of('.');
		if( std::string::npos == l_Idx ) return false;
		l_Ext = a_Filepath.substr(l_Idx + 1);
	}
	
	if( "xlsx" == l_Ext ) m_pLoader = new XlsxLoader();
	else if( "xls" == l_Ext ) m_pLoader = new XlsLoader();

	if( nullptr == m_pLoader ) return false;

	return m_pLoader->load(a_Filepath);
}

bool ExcelAdapter::setSheet(std::string a_SheetName)
{
	assert(nullptr != m_pLoader);
	return m_pLoader->setSheet(a_SheetName);
}

std::string ExcelAdapter::getCell(unsigned int a_Row, unsigned int a_Column)
{
	assert(nullptr != m_pLoader);
	return m_pLoader->getCell(a_Row, a_Column);
}
#pragma endregion