// Converter.h
//
// 2017/10/29 Ian Wu/Real0000
//

#ifndef _CONVERTER_H_
#define _CONVERTER_H_

#include <map>

#include "boost/property_tree/ptree.hpp"
#include "rapidjson/document.h"
#include "ExcelAdapter.h"

class Converter
{
public:
	Converter();
	virtual ~Converter();

	void clear();

	void convert(std::string a_SettingFile, bool a_bPretty);

private:
	void initSettings(boost::property_tree::ptree &a_SettingRoot);
	void output(boost::property_tree::ptree &a_FileRoot, bool a_bPretty);

	enum CellType
	{
		CELL_INT = 0,
		CELL_FLOAT,
		CELL_TEXT
	};
	enum LinkType
	{
		LINK_NONE = 0,
		LINK_OBJ,		// obj
		LINK_DICT,		// dict
		LINK_ARRAY,		// array
	};
	typedef std::pair<unsigned int, CellType> ColumnInfo;
	struct Table
	{
		std::map<std::string, ColumnInfo> m_Column;
		std::string m_Sheetname;
		std::string m_Tablename;
		unsigned int m_StartRow, m_StartColumn, m_NumColumn, m_NumRow;
	};
	struct TableNode
	{
		TableNode()
			: m_LinkType(LINK_NONE)
			, m_KeyColumn(0)
			, m_pRefTable(nullptr), m_pChildRefTable(nullptr)
			, m_Type(CELL_TEXT), m_Name(""), m_Column(0)
		{}
		~TableNode()
		{
			for( unsigned int i=0 ; i<m_Children.size() ; ++i ) delete m_Children[i];
			m_Children.clear();
		}

		// node
		LinkType m_LinkType;
		unsigned int m_KeyColumn;
		unsigned int m_LinkColumn;
		Table *m_pRefTable, *m_pChildRefTable;
		std::vector<TableNode *> m_Children;

		// data
		CellType m_Type;
		std::string m_Name;
		unsigned int m_Column;

	};
	void initOutputNode(TableNode *a_pOwner, boost::property_tree::ptree &a_Node, Table *a_pCurrTable);
	void initOutputBuffer(rapidjson::Document &a_Root, rapidjson::Value &a_Owner, TableNode *a_pNode, unsigned int a_Row);
	void initUE4Buffer(std::string &a_HeaderOutput, std::string &a_CppOutput, std::string a_ClassName, std::string a_ValueName, CellType a_KeyType, CellType a_ValueType);
	void initUE4Buffer(std::string &a_HeaderOutput, std::string &a_CppOutput, std::string a_ClassName, TableNode *a_pNode);
	std::string cellType2UE4Type(CellType a_Type);
	std::string cellType2UE4Value(CellType a_Type, std::string a_TargetStr, std::string a_SrcStr, bool a_bSrcIsJson);

	std::map<std::string, Table*> m_Tables;
	ExcelAdapter m_Source;
	std::string m_UE4ProjectName;// for ue4 output
};

#endif