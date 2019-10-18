// Converter.h
//
// 2017/10/29 Ian Wu/Real0000
//
#include <direct.h>
#include <deque>
#include <set>

#include "Converter.h"

#include "boost/property_tree/xml_parser.hpp"

#include "rapidjson/memorybuffer.h"
#include "rapidjson/prettywriter.h"
#include "rapidjson/writer.h"


#pragma region Converter
//
// Converter
//
Converter::Converter()
	: m_UE4ProjectName("")
{
}

Converter::~Converter()
{
	clear();
}

void Converter::clear()
{
	m_Source.clear();
	for( auto it = m_Tables.begin() ; it != m_Tables.end() ; ++it ) delete it->second;
	m_Tables.clear();
}

void Converter::convert(std::string a_SettingFile, bool a_bPretty)
{
	clear();
	
	boost::property_tree::ptree l_Root;
	boost::property_tree::xml_parser::read_xml(a_SettingFile, l_Root);
	if( l_Root.empty() ) return ;

	{// init source file
		std::string l_ExcelFile(l_Root.get("root.<xmlattr>.filename", ""));
		std::string l_Chdir(l_Root.get("root.<xmlattr>.rootDir", ""));
		m_UE4ProjectName = l_Root.get("root.<xmlattr>.ue4ProjectName", "");
		if( l_ExcelFile.empty() ) return;

		_chdir(l_Chdir.c_str());
		if( !m_Source.load(l_ExcelFile) ) return;
	}

	initSettings(l_Root.get_child("root.Tables"));
	boost::property_tree::ptree &l_Outputs = l_Root.get_child("root.Outputs");
	for( auto it = l_Outputs.begin() ; it != l_Outputs.end() ; ++it )
	{
		if( "<xmlattr>" == it->first ) continue;
		output(it->second, a_bPretty);
	}
}

void Converter::initSettings(boost::property_tree::ptree &a_SettingRoot)
{
	for( auto it = a_SettingRoot.begin() ; it != a_SettingRoot.end() ; ++it )
	{
		if( "<xmlattr>" == it->first ) continue;

		std::string l_SheetName(it->second.get("<xmlattr>.sheet", ""));
		if( !m_Source.setSheet(l_SheetName.c_str()) ) continue;

		unsigned int l_DescRow = it->second.get("<xmlattr>.descRow", 0);
		unsigned int l_StartCol = 0;
		{
			std::string l_ColStr(it->second.get("<xmlattr>.startColumn", "0"));
			if( isdigit(l_ColStr[0]) ) l_StartCol = atoi(l_ColStr.c_str());
			else
			{
				auto l_ColIt = std::find_if(l_ColStr.begin(), l_ColStr.end(), [] (char a_Char){ return !isalpha(a_Char) && '\0' != a_Char; });
				if( l_ColIt != l_ColStr.end() )
				{
					printf("invalid column %s in table %s\n", l_ColStr.c_str(), l_SheetName.c_str());
					continue;
				}

				for( unsigned int i=0 ; i<l_ColStr.size() ; ++i ) l_ColStr[i] = tolower(l_ColStr[i]);
				for( unsigned int i=0 ; i<l_ColStr.size() ; ++i )
				{
					char l_Char = l_ColStr[l_ColStr.size() - i - 1];
					l_StartCol += (l_Char - 'a') * (unsigned int)pow(26, i);
				}
			}
		}

		if( m_Source.getCell(l_DescRow, l_StartCol).empty() ) continue;

		std::string l_TableName(it->second.get("<xmlattr>.name", ""));
		if( l_TableName.empty() || m_Tables.end() != m_Tables.find(l_TableName) ) continue;

		Table *l_pNewTable = new Table();
		m_Tables.insert(std::make_pair(l_TableName, l_pNewTable));
		l_pNewTable->m_Sheetname = l_SheetName;
		l_pNewTable->m_Tablename = l_TableName;
		l_pNewTable->m_StartRow = l_DescRow + 1;
		l_pNewTable->m_StartColumn = l_StartCol;

		unsigned int l_Col = l_StartCol;
		std::string l_DescName("");
		std::map<std::string, unsigned int> l_DescMap;
		l_DescMap.clear();
		while( !(l_DescName = m_Source.getCell(l_DescRow, l_Col)).empty() )
		{
			l_DescMap[l_DescName] = l_Col;
			++l_Col;
		}
		l_pNewTable->m_NumColumn = l_Col - l_StartCol;

		l_pNewTable->m_NumRow = 0;
		while( !m_Source.getCell(l_DescRow + l_pNewTable->m_NumRow + 1, l_StartCol).empty() ) ++l_pNewTable->m_NumRow;

		for( auto l_TypeIt = it->second.begin() ; l_TypeIt != it->second.end() ; ++l_TypeIt )
		{
			if( "<xmlattr>" == l_TypeIt->first ) continue;

			auto l_Element = l_DescMap.find(l_TypeIt->second.get("<xmlattr>.name", ""));
			if( l_DescMap.end() == l_Element ) continue;

			CellType l_CellType = CELL_TEXT;
			std::string l_TypeStr(l_TypeIt->second.get("<xmlattr>.type", ""));
			if( "int" == l_TypeStr ) l_CellType = CELL_INT;
			else if( "float" == l_TypeStr ) l_CellType = CELL_FLOAT;

			l_pNewTable->m_Column.insert(std::make_pair(l_Element->first, std::make_pair(l_Element->second, l_CellType)));
		}
	}
}

void Converter::output(boost::property_tree::ptree &a_FileRoot, bool a_bPretty)
{
	TableNode *l_pRoot = nullptr;
	
	Table *l_pRootTable = nullptr;
	bool l_bSimple = false;
	std::string l_UE4Class("");
	{// init root
		std::string l_TableName(a_FileRoot.get("<xmlattr>.rootTable", ""));
		auto l_TableIt = m_Tables.find(l_TableName);
		if( m_Tables.end() == l_TableIt ) return;
		l_pRootTable = l_TableIt->second;

		std::string l_KeyID(a_FileRoot.get("<xmlattr>.key", ""));
		auto l_KeyIt = l_TableIt->second->m_Column.find(l_KeyID);
		if( l_TableIt->second->m_Column.end() == l_KeyIt ) return;

		l_bSimple = a_FileRoot.get("<xmlattr>.simple", 0) != 0;
		l_UE4Class = a_FileRoot.get("<xmlattr>.ue4Class", "");

		if( !l_bSimple )
		{
			l_pRoot = new TableNode();
			l_pRoot->m_pRefTable = l_pRootTable;
			l_pRoot->m_KeyColumn = l_KeyIt->second.first;
			l_pRoot->m_Type = l_KeyIt->second.second;
		}
	}
	
	std::string l_UE4Header(""), l_UE4Cpp("");
	rapidjson::Document l_FileBuffer(rapidjson::kObjectType);
	if( l_bSimple )
	{
		ColumnInfo *l_pKeyColumn = nullptr;
		ColumnInfo *l_pDataColumn = nullptr;
		std::string l_KeyName(a_FileRoot.get("<xmlattr>.key", ""));
		std::string l_DataName(a_FileRoot.get("<xmlattr>.data", ""));
		{
			auto l_KeyIt = l_pRootTable->m_Column.find(l_KeyName);
			if( l_pRootTable->m_Column.end() != l_KeyIt ) l_pKeyColumn = &l_KeyIt->second;

			auto l_DataIt = l_pRootTable->m_Column.find(l_DataName);
			if( l_pRootTable->m_Column.end() != l_DataIt ) l_pDataColumn = &l_DataIt->second;
		}

		if( nullptr == l_pKeyColumn || nullptr == l_pDataColumn ) return;

		std::function<void(rapidjson::Document&, std::string&, std::string&)> l_pAddFunc = nullptr;
		switch( l_pDataColumn->second )
		{
			case CELL_INT:
				l_pAddFunc = [=](rapidjson::Document &a_Target, std::string &a_Key, std::string &a_Value)
				{
					auto &l_Allocator = a_Target.GetAllocator();

					rapidjson::Value l_Key(a_Key.c_str(), l_Allocator);
					a_Target.AddMember(l_Key, atoi(a_Value.c_str()), l_Allocator);
				};
				break;

			case CELL_FLOAT:
				l_pAddFunc = [=](rapidjson::Document &a_Target, std::string &a_Key, std::string &a_Value)
				{
					auto &l_Allocator = a_Target.GetAllocator();

					rapidjson::Value l_Key(a_Key.c_str(), l_Allocator);
					a_Target.AddMember(l_Key, atof(a_Value.c_str()), l_Allocator);
				};
				break;

			case CELL_TEXT:
				l_pAddFunc = [=](rapidjson::Document &a_Target, std::string &a_Key, std::string &a_Value)
				{
					auto &l_Allocator = a_Target.GetAllocator();

					rapidjson::Value l_Key(a_Key.c_str(), l_Allocator);
					rapidjson::Value l_Value(a_Value.c_str(), l_Allocator);
					a_Target.AddMember(l_Key, l_Value, l_Allocator);
				};
				break;

			default:break;
		}

		m_Source.setSheet(l_pRootTable->m_Sheetname);
		for( unsigned int r=0 ; r<l_pRootTable->m_NumRow ; ++r )
		{
			l_pAddFunc(l_FileBuffer, m_Source.getCell(r + l_pRootTable->m_StartRow, l_pKeyColumn->first), m_Source.getCell(r + l_pRootTable->m_StartRow, l_pDataColumn->first));
		}

		if( !l_UE4Class.empty() )
		{
			std::string l_UE4ClassName(l_UE4Class.substr(l_UE4Class.find_last_of('/') + 1));
			initUE4Buffer(l_UE4Header, l_UE4Cpp, l_UE4ClassName, l_DataName, l_pKeyColumn->second, l_pDataColumn->second);
		}
	}
	else
	{
		// init tree
		initOutputNode(l_pRoot, a_FileRoot, l_pRootTable);

		// init buffer
		auto &l_Allocator = l_FileBuffer.GetAllocator();
		m_Source.setSheet(l_pRootTable->m_Sheetname);
		for( unsigned int r=0 ; r<l_pRootTable->m_NumRow ; ++r )
		{
			rapidjson::Value l_Obj;
			l_Obj.SetObject();

			initOutputBuffer(l_FileBuffer, l_Obj, l_pRoot, r + l_pRootTable->m_StartRow);
			m_Source.setSheet(l_pRootTable->m_Sheetname);

			rapidjson::Value l_Key(m_Source.getCell(r + l_pRootTable->m_StartRow, l_pRoot->m_KeyColumn).c_str(), l_Allocator);
			l_FileBuffer.AddMember(l_Key, l_Obj, l_Allocator);
		}
		
		if( !l_UE4Class.empty() )
		{
			std::string l_UE4ClassName(l_UE4Class.substr(l_UE4Class.find_last_of('/') + 1));
			initUE4Buffer(l_UE4Header, l_UE4Cpp, l_UE4ClassName, l_pRoot);
		}
	}

	// write buffer
	rapidjson::StringBuffer l_Output;
	if( a_bPretty )
	{
		rapidjson::PrettyWriter<rapidjson::StringBuffer> l_Writer(l_Output);
		l_FileBuffer.Accept(l_Writer);
	}
	else
	{
		rapidjson::Writer<rapidjson::StringBuffer> l_Writer(l_Output);
		l_FileBuffer.Accept(l_Writer);
	}

	FILE *l_pFile = fopen(a_FileRoot.get("<xmlattr>.file", "").c_str(), "wb");
	if( nullptr != l_pFile )
	{
		fwrite(l_Output.GetString(), sizeof(char), l_Output.GetSize(), l_pFile);
		fclose(l_pFile);
	}

	// write ue4 class
	if( !l_UE4Class.empty() )
	{
		l_pFile = fopen((l_UE4Class + ".h").c_str(), "wb");
		fwrite(l_UE4Header.c_str(), sizeof(char), l_UE4Header.length(), l_pFile);
		fclose(l_pFile);

		l_pFile = fopen((l_UE4Class + ".cpp").c_str(), "wb");
		fwrite(l_UE4Cpp.c_str(), sizeof(char), l_UE4Cpp.length(), l_pFile);
		fclose(l_pFile);
	}

	// clean up
	delete l_pRoot;
	l_pRoot = nullptr;
}

void Converter::initOutputNode(TableNode *a_pOwner, boost::property_tree::ptree &a_Node, Table *a_pCurrTable)
{
	for( auto it = a_Node.begin() ; it != a_Node.end() ; ++it )
	{
		if( "<xmlattr>" == it->first ) continue;
	
		boost::property_tree::ptree &l_AttrNode = it->second.get_child("<xmlattr>");

		std::string l_ColumnName(l_AttrNode.get("name", ""));
		auto l_KeyIt = a_pCurrTable->m_Column.find(l_ColumnName);
		if( a_pCurrTable->m_Column.end() == l_KeyIt ) continue;

		// init data part
		TableNode *l_pNewNode = new TableNode();
		a_pOwner->m_Children.push_back(l_pNewNode);

		l_pNewNode->m_Name = l_ColumnName;
		l_pNewNode->m_Type = l_KeyIt->second.second;
		l_pNewNode->m_pRefTable = a_pCurrTable;
		l_pNewNode->m_Column = l_KeyIt->second.first;

		std::string l_LinkType(l_AttrNode.get("linkType", ""));
		if( "dict" == l_LinkType ) l_pNewNode->m_LinkType = LINK_DICT;
		else if( "array" == l_LinkType ) l_pNewNode->m_LinkType = LINK_ARRAY;
		else if( "obj" == l_LinkType ) l_pNewNode->m_LinkType = LINK_OBJ;
		else continue;

		std::string l_TableName(l_AttrNode.get("table", ""));
		auto l_TableIt = m_Tables.find(l_TableName);
		if( m_Tables.end() == l_TableIt ) continue;

		auto l_LinkIt = l_TableIt->second->m_Column.find(l_pNewNode->m_Name);
		if( l_TableIt->second->m_Column.end() == l_LinkIt ) continue;
		l_pNewNode->m_LinkColumn = l_LinkIt->second.first;
		l_pNewNode->m_pChildRefTable = l_TableIt->second;

		if( LINK_ARRAY != l_pNewNode->m_LinkType )
		{
			std::string l_SubkeyName(l_AttrNode.get("key", ""));
			auto l_SubKeyIt = l_TableIt->second->m_Column.find(l_SubkeyName);
			if( l_TableIt->second->m_Column.end() == l_SubKeyIt ) continue;

			l_pNewNode->m_KeyColumn = l_SubKeyIt->second.first;
		}
		initOutputNode(l_pNewNode, it->second, l_TableIt->second);
	}
}

void Converter::initOutputBuffer(rapidjson::Document &a_Root, rapidjson::Value &a_Owner, TableNode *a_pNode, unsigned int a_Row)
{
	auto &l_Allocator = a_Root.GetAllocator();
	for( unsigned int i=0 ; i<a_pNode->m_Children.size() ; ++i )
	{
		TableNode *l_pCurrChild = a_pNode->m_Children[i];
		rapidjson::Value l_Key(l_pCurrChild->m_Name.c_str(), l_Allocator);
		if( l_pCurrChild->m_Children.empty() )
		{	
			m_Source.setSheet(l_pCurrChild->m_pRefTable->m_Sheetname);
			switch( l_pCurrChild->m_Type )
			{
				case CELL_FLOAT:{
					double l_Val = atof(m_Source.getCell(a_Row, l_pCurrChild->m_Column).c_str());
					a_Owner.AddMember(l_Key, l_Val, l_Allocator);
					}break;

				case CELL_INT:{
					int l_Val = atoi(m_Source.getCell(a_Row, l_pCurrChild->m_Column).c_str());
					a_Owner.AddMember(l_Key, l_Val, l_Allocator);
					}break;
					
				// case CELL_TEXT:
				default:{
					rapidjson::Value l_Val(m_Source.getCell(a_Row, l_pCurrChild->m_Column).c_str(), l_Allocator);
					a_Owner.AddMember(l_Key, l_Val, l_Allocator);
					}break;
			}
		}
		else
		{
			std::string l_KeyValue(m_Source.getCell(a_Row, l_pCurrChild->m_Column));
			Table *l_pChildTable = l_pCurrChild->m_pChildRefTable;
			m_Source.setSheet(l_pChildTable->m_Sheetname);
			unsigned int l_CheckColumn = l_pCurrChild->m_LinkColumn;
			switch( l_pCurrChild->m_LinkType )
			{
				case LINK_OBJ:{
					rapidjson::Value l_Obj(rapidjson::kObjectType);
					for( unsigned int r=0 ; r<l_pChildTable->m_NumRow ; ++r )
					{
						if( m_Source.getCell(r + l_pChildTable->m_StartRow, l_CheckColumn) == l_KeyValue )
						{
							initOutputBuffer(a_Root, l_Obj, l_pCurrChild, r + l_pChildTable->m_StartRow);
							break;
						}
					}
					a_Owner.AddMember(l_Key, l_Obj, l_Allocator);
					}break;

				case LINK_DICT:{
					rapidjson::Value l_Dict(rapidjson::kObjectType);
					for( unsigned int r=0 ; r<l_pChildTable->m_NumRow ; ++r )
					{
						if( m_Source.getCell(r + l_pChildTable->m_StartRow, l_CheckColumn) == l_KeyValue )
						{
							rapidjson::Value l_Obj(rapidjson::kObjectType);
							initOutputBuffer(a_Root, l_Obj, l_pCurrChild, r + l_pChildTable->m_StartRow);

							std::string l_SubKeyStr(m_Source.getCell(r + l_pChildTable->m_StartRow, l_pCurrChild->m_KeyColumn));
							rapidjson::Value l_SubKey(l_SubKeyStr.c_str(), l_Allocator);
							l_Dict.AddMember(l_SubKey, l_Obj, l_Allocator);
						}
					}
					a_Owner.AddMember(l_Key, l_Dict, l_Allocator);
					}break;

				case LINK_ARRAY:{
					rapidjson::Value l_Array(rapidjson::kArrayType);
					for( unsigned int r=0 ; r<l_pChildTable->m_NumRow ; ++r )
					{
						if( m_Source.getCell(r + l_pChildTable->m_StartRow, l_CheckColumn) == l_KeyValue )
						{
							rapidjson::Value l_Obj(rapidjson::kObjectType);
							initOutputBuffer(a_Root, l_Obj, l_pCurrChild, r + l_pChildTable->m_StartRow);
							l_Array.PushBack(l_Obj, l_Allocator);
						}
					}

					a_Owner.AddMember(l_Key, l_Array, l_Allocator);
					}break;

				default:break;
			}
			m_Source.setSheet(l_pCurrChild->m_pRefTable->m_Sheetname);
		}
	}
}

void Converter::initUE4Buffer(std::string &a_HeaderOutput, std::string &a_CppOutput, std::string a_ClassName, std::string a_ValueName, Converter::CellType a_KeyType, Converter::CellType a_ValueType)
{
	// setup header content
	a_HeaderOutput = 
		"// generate by Excel2Json\n\n"
		"#pragma once\n\n"
		"#include \"CoreMinimal.h\"\n"
		"#include \"Engine/DataAsset.h\"\n"
		"#include \"" + a_ClassName + ".generated.h\"\n\n"
		"UCLASS()\n"
		"class U" + a_ClassName + " : public UDataAsset\n"
		"{\n"
		"\tGENERATED_BODY()\n\n"
		"public:\n"
		"\tvoid parseTable(FString a_Filename);\n\n"
		"\tUPROPERTY(BlueprintReadOnly, Category=\"DataTable\")\n";
			
	char l_Buff[256];
	snprintf(l_Buff, 256, "\tTMap<%s, %s> m_%s;\n", cellType2UE4Type(a_KeyType).c_str(), cellType2UE4Type(a_ValueType).c_str(), a_ValueName.c_str());
	a_HeaderOutput += l_Buff;

	a_HeaderOutput += "};\n";

	// setup cpp content
	a_CppOutput =
		"// generate by Excel2Json\n\n"
		"#include \"" + a_ClassName + ".h\"\n"
		"#include \"" + m_UE4ProjectName + ".h\"\n"
		"#include \"JsonUtilities.h\"\n\n"
		"void U" + a_ClassName + "::parseTable(FString a_Filename)\n"
		"{\n"
		"\tFString l_FileString;\n"
		"\tFFileHelper::LoadFileToString(l_FileString, *a_Filename);\n"
		"\tTSharedPtr<FJsonObject> l_JsonObject = MakeShareable(new FJsonObject());\n"
		"\tTSharedRef<TJsonReader<> > l_JsonReader = TJsonReaderFactory<>::Create(l_FileString);\n"
		"\tif( !FJsonSerializer::Deserialize(l_JsonReader, l_JsonObject) || !l_JsonObject.IsValid() ) return;\n\n"
		"\tTMap< FString, TSharedPtr<FJsonValue> > &l_Root = l_JsonObject->Values;\n"
		"\tfor( const TPair<FString, TSharedPtr<FJsonValue> > &l_Pair : l_Root )\n"
		"\t{\n"
		"\t\t" + cellType2UE4Value(a_KeyType, cellType2UE4Type(a_KeyType) + " l_Key", "l_Pair.Key", false) + "\n"
		"\t\t" + cellType2UE4Value(a_ValueType, cellType2UE4Type(a_ValueType) + " l_Value", "l_Pair.Value", true) + "\n"
		"\t\tm_" + a_ValueName + ".Add(l_Key, l_Value);\n"
		"\t}\n"
		"}";
}

void Converter::initUE4Buffer(std::string &a_HeaderOutput, std::string &a_CppOutput, std::string a_ClassName, TableNode *a_pNode)
{
	std::map<Table *, std::set<TableNode *> > l_Tables;
	{
		std::vector<TableNode *> l_Nodes(1, a_pNode);
		for( unsigned int i=0 ; i<l_Nodes.size() ; ++i )
		{
			for( unsigned int j=0 ; j<l_Nodes[i]->m_Children.size() ; ++j ) l_Nodes.push_back(l_Nodes[i]->m_Children[j]);
			if( l_Nodes[i]->m_Name.empty() ) continue;

			auto it = l_Tables.find(l_Nodes[i]->m_pRefTable);
			if( l_Tables.end() == it )
			{
				std::set<TableNode *> l_NewSet;
				l_NewSet.insert(l_Nodes[i]);
				l_Tables.insert(std::make_pair(l_Nodes[i]->m_pRefTable, l_NewSet));
			}
			else it->second.insert(l_Nodes[i]);
		}
	}

	// setup header content
	a_HeaderOutput = 
		"// generate by Excel2Json\n\n"
		"#pragma once\n\n"
		"#include \"CoreMinimal.h\"\n"
		"#include \"Engine/DataAsset.h\"\n"
		"#include \"" + a_ClassName + ".generated.h\"\n\n";
	for( auto it = l_Tables.begin() ; it != l_Tables.end() ; ++it )
	{
		a_HeaderOutput += "class U" + it->first->m_Tablename + ";\n";
	}
	a_HeaderOutput += "class U" + a_ClassName + "Parser;\n\n";

	for( auto it = l_Tables.begin() ; it != l_Tables.end() ; ++it )
	{
		a_HeaderOutput +=
			"UCLASS()\n"
			"class U" + it->first->m_Tablename + " : public UDataAsset\n"
			"{\n"
			"\tGENERATED_BODY()\n\n"
			"public:\n"
			"\tvoid parseTable(TSharedPtr<FJsonObject> a_pNode);\n\n";
			"private:\n";

		std::set<TableNode *> &l_NodeSet = it->second;
		for( auto l_ParamIt = l_NodeSet.begin() ; l_ParamIt != l_NodeSet.end() ; ++l_ParamIt )
		{
			a_HeaderOutput += "\tUPROPERTY(BlueprintReadOnly, Category=\"DataTable\")\n";
			switch( (*l_ParamIt)->m_LinkType )
			{
				case LINK_OBJ:{
					a_HeaderOutput += "\tU" + (*l_ParamIt)->m_pChildRefTable->m_Tablename + "* m_p" + (*l_ParamIt)->m_Name;
					}break;

				case LINK_DICT:{
					a_HeaderOutput += "\tTMap<" + cellType2UE4Type((*l_ParamIt)->m_Type) + ", U" + (*l_ParamIt)->m_pChildRefTable->m_Tablename + "*> m_" + (*l_ParamIt)->m_Name;
					}break;

				case LINK_ARRAY:{
					a_HeaderOutput += "\tTArray<U" + (*l_ParamIt)->m_pChildRefTable->m_Tablename + "*> m_" + (*l_ParamIt)->m_Name;
					}break;

				//case LINK_NONE:
				default:{
					a_HeaderOutput += "\t" + cellType2UE4Type((*l_ParamIt)->m_Type) + " m_" + (*l_ParamIt)->m_Name;
					}break;
			}
			a_HeaderOutput += ";\n\n";
		}
		a_HeaderOutput += "};\n\n";
	}
	
	a_HeaderOutput +=
		"UCLASS()\n"
		"class U" + a_ClassName + " : public UDataAsset\n"
		"{\n"
		"\tGENERATED_BODY()\n\n"
		"public:\n"
		"\tvoid parseTable(FString a_Filename);\n\n"
		"\tUPROPERTY(BlueprintReadOnly, Category=\"DataTable\")\n";

	char l_Buff[256];
	snprintf(l_Buff, 256, "\tTMap<%s, U%s*> m_%s;\n", cellType2UE4Type(a_pNode->m_Type).c_str(), a_pNode->m_pRefTable->m_Tablename.c_str(), a_pNode->m_pRefTable->m_Tablename.c_str());
	a_HeaderOutput += l_Buff;

	a_HeaderOutput += "};\n";

	// setup cpp content
	a_CppOutput =
		"// generate by Excel2Json\n\n"
		"#include \"" + a_ClassName + ".h\"\n"
		"#include \"" + m_UE4ProjectName + ".h\"\n"
		"#include \"JsonUtilities.h\"\n\n";
	for( auto it = l_Tables.begin() ; it != l_Tables.end() ; ++it )
	{
		a_CppOutput +=
			"void U" + it->first->m_Tablename + "::parseTable(TSharedPtr<FJsonObject> a_pNode)\n"
			"{\n";
			
		std::set<TableNode *> &l_NodeSet = it->second;
		for( auto l_ParamIt = l_NodeSet.begin() ; l_ParamIt != l_NodeSet.end() ; ++l_ParamIt )
		{
			switch( (*l_ParamIt)->m_LinkType )
			{
				case LINK_OBJ:{
					a_CppOutput +=
						"\tm_p" + (*l_ParamIt)->m_Name + " = NewObject<U" + (*l_ParamIt)->m_pChildRefTable->m_Tablename + ">();\n"
						"\tm_p" + (*l_ParamIt)->m_Name + "->parseTable(" + "a_pNode->GetObjectField("+ (*l_ParamIt)->m_Name +")" +");\n";
					}break;

				case LINK_DICT:{
					std::string l_ChildTypeStr = (*l_ParamIt)->m_pChildRefTable->m_Tablename;
					a_CppOutput +=
						"\n\t{\n"
						"\t\tconst TSharedPtr<FJsonObject> l_Obj = a_pNode->GetObjectField(\"" + l_ChildTypeStr + "\")\n"
						"\t\tfor( const TPair<FString, TSharedPtr<FJsonValue> > &l_Pair : l_Obj->Values )\n"
						"\t\t{\n"
						"\t\t\t" + cellType2UE4Value((*l_ParamIt)->m_Type, cellType2UE4Type((*l_ParamIt)->m_Type) + " l_Key", "l_Pair.Key", false) + "\n"
						"\t\t\tU" + l_ChildTypeStr + " *l_pNewObj = NewObject<U" + l_ChildTypeStr + ">();\n"
						"\t\t\tl_pNewObj->parseTable(l_Pair.Value->AsObject());\n"
						"\t\t\tm_" + l_ChildTypeStr + ".Add(l_Key, l_pNewObj);\n"
						"\t\t}\n"
						"\t}\n\n";
					}break;

				case LINK_ARRAY:{
					std::string l_ChildTypeStr = (*l_ParamIt)->m_pChildRefTable->m_Tablename;
					a_CppOutput +=
						"\n\t{\n"
						"\t\tconst TArray<TSharedPtr<FJsonValue> > l_Array = a_pNode->GetArrayField(\"" + l_ChildTypeStr + "\");\n"
						"\t\tfor( int32 i=0 ; i<l_Array.Num() ; ++i )\n"
						"\t\t{\n"
						"\t\t\tU" + l_ChildTypeStr + " *l_pNewItem = NewObject<U" + l_ChildTypeStr + ">();\n"
						"\t\t\tl_pNewItem->parseTable(l_Array[i]->AsObject());\n"
						"\t\t\tm_" + (*l_ParamIt)->m_Name + ".Push(l_pNewItem);\n"
						"\t\t}\n"
						"\t}\n\n";
					}break;

				//case LINK_NONE:
				default:{
					switch( (*l_ParamIt)->m_Type )
					{
						case CELL_INT:
							a_CppOutput += "\tm_" + (*l_ParamIt)->m_Name + " = " + "a_pNode->GetIntegerField(\"" + (*l_ParamIt)->m_Name +"\");\n";
							break;

						case CELL_FLOAT:
							a_CppOutput += "\tm_" + (*l_ParamIt)->m_Name + " = " + "a_pNode->GetNumberField(\"" + (*l_ParamIt)->m_Name +"\");\n";
							break;

						case CELL_TEXT:
							a_CppOutput += "\tm_" + (*l_ParamIt)->m_Name + " = " + "a_pNode->GetStringField(\"" + (*l_ParamIt)->m_Name +"\");\n";
							break;

						default:
							assert(false && "invalid key cell type");
							break;
					}
					}break;
			}
		}
		a_CppOutput += "}\n\n";
	}

	a_CppOutput +=
		"void U" + a_ClassName + "::parseTable(FString a_Filename)\n"
		"{\n"
		"\tFString l_FileString;\n"
		"\tFFileHelper::LoadFileToString(l_FileString, *a_Filename);\n"
		"\tTSharedPtr<FJsonObject> l_JsonObject = MakeShareable(new FJsonObject());\n"
		"\tTSharedRef<TJsonReader<> > l_JsonReader = TJsonReaderFactory<>::Create(l_FileString);\n"
		"\tif( !FJsonSerializer::Deserialize(l_JsonReader, l_JsonObject) || !l_JsonObject.IsValid() ) return;\n\n"
		"\tTMap< FString, TSharedPtr<FJsonValue> > &l_Root = l_JsonObject->Values;\n"
		"\tfor( const TPair<FString, TSharedPtr<FJsonValue> > &l_Pair : l_Root )\n"
		"\t{\n"
		"\t\t" + cellType2UE4Value(a_pNode->m_Type, cellType2UE4Type(a_pNode->m_Type) + " l_Key", "l_Pair.Key", false) + "\n"
		"\t\tU" + a_pNode->m_pRefTable->m_Tablename + " *l_pNewObj = NewObject<U" + a_pNode->m_pRefTable->m_Tablename + ">();\n"
		"\t\tl_pNewObj->parseTable(l_Pair.Value->AsObject());\n"
		"\t\tm_" + a_pNode->m_pRefTable->m_Tablename + ".Add(l_Key, l_pNewObj);\n"
		"\t}\n"
		"}";
}

std::string Converter::cellType2UE4Type(CellType a_Type)
{
	switch( a_Type )
	{
		case CELL_INT:	return "int32";
		case CELL_FLOAT:return "float";
		case CELL_TEXT:	return "FString";
		default: assert(false && "unknown cell type");
	}
	return "";
}

std::string Converter::cellType2UE4Value(CellType a_Type, std::string a_TargetStr, std::string a_SrcStr, bool a_bSrcIsJson)
{
	switch( a_Type )
	{
		case CELL_INT:	return a_TargetStr + " = " + (a_bSrcIsJson ? "std::floor(" + a_SrcStr + "->AsNumber())" : "FCString::Atoi(*" + a_SrcStr + ");");
		case CELL_FLOAT:return a_TargetStr + " = " + (a_bSrcIsJson ? a_SrcStr + "->AsNumber();" : "FCString::Atof(*" + a_SrcStr + ");");
		case CELL_TEXT:	return a_TargetStr + " = " + a_SrcStr + (a_bSrcIsJson ? "->AsString();" : ";");
		default:
			assert(false && "invalid key cell type");
			break;
	}
	return "";
}
#pragma endregion