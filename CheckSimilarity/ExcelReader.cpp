#include "stdafx.h"
#include "ExcelReader.h"
#include <codecvt>

#include <regex>							//正则表达式

ExcelReader::ExcelReader()
{
	curRow = 1;
	isOpen = false;
}

ExcelReader::~ExcelReader()
{
}

void ExcelReader::addXlsxFile(const std::string & filename)
{
	fileNames.push_back(filename);
}

void ExcelReader::loadXlsxFile(const std::string & filename)
{
	if (isOpen) {
		//重置状态
		delete wb;
		curRow = 1;
		selColumns.clear();
		fileNames.clear();
		//isOpen = false;
	}

	wb = new xlnt::workbook();	
	wb->load(filename);		//为了支持中文路径,传入的路径字符串已经是宽字符编码
	ws = wb->active_sheet();
	
	//统计总行数
	//最大行数减一，去掉列名
	auto rows = ws.rows(false);
	maxRow = rows.length() - 1;

	isOpen = true;

}

void ExcelReader::loadXlsxFile(const std::string & pattern, const std::string & partOfSpeech)
{
	std::regex re(pattern);

	for (auto tmp : fileNames) {
		bool ret = std::regex_match(tmp, re);



		if (ret)
			break;
		else 
			fprintf(stderr, "%s, can not match\n", tmp.c_str());

	}


}

bool ExcelReader::isOpenFile()
{
	return isOpen;
}

// 如果已达到最后一行，则返回false
bool ExcelReader::nextWord()
{
	//最大行数减一，去掉列名

	if (curRow < maxRow) {
		curRow++;
		return true;
	}
	else
		return false;
}

void ExcelReader::selectColumn(const std::string & columnName)
{
	auto columns = ws.columns(false);
	for (auto column : columns) {
		std::string str = column[0].to_string();

		//使用xLnt读取xlsx文件，返回值均为utf-8编码
		//故str中实际存储的是utf-8编码的字符串

		if (columnName == str) {
			selColumns.push_back(make_pair(columnName, column));
			return;
		}

	}

}

std::string ExcelReader::getCurCellValueInColumn(const std::string & columnName)
{
	for (auto pair : selColumns) {
		if (pair.first == columnName) {
			return pair.second[curRow].to_string();
		}
	}

	return std::string("none");
}



