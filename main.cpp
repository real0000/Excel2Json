// main.cpp
//
// 2017/10/29 Ian Wu/Real0000
//

#include <iostream>
#include <string>
#include <direct.h>

#include "Converter.h"

//
// command : Excel2Json -f <setting file path> [-pretty]
//

int main(int a_Argc, char *a_Argv[])
{
	bool l_bPretty = false;
	std::string l_File("");
	for( int i=1 ; i<a_Argc ; ++i )
	{
		if( 0 == strcmp(a_Argv[i], "-f") )
		{
			if( i == a_Argc - 1 ) return -1;
			l_File = a_Argv[i+1];
		}
		else if( 0 == strcmp(a_Argv[i], "-pretty") )
		{
			l_bPretty = true;
		}
	}
	
	if( l_File.empty() ) return -1;

	Converter a_Loader;
	a_Loader.convert(l_File, l_bPretty);

	return 0;
}