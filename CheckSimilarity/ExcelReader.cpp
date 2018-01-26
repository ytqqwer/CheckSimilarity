#include "stdafx.h"
#include "ExcelReader.h"
#include <codecvt>

#include <regex>							//正则表达式

ExcelReader::ExcelReader()
{
	maxRow = curRow = 1;
	curWorkbookIndex = 0;
	existingFile = false;
}

ExcelReader::~ExcelReader()
{
}

void ExcelReader::addXlsxFileName(const std::string & filename)
{
	fileNames.push_back(filename);
}

void ExcelReader::clear()
{
	//if (isOpen) {
		////重置状态
		//curRow = 1;
		//selColumns.clear();
		//fileNames.clear();
		//isOpen = false;
	//}

	//wb = new xlnt::workbook();
	//wb->load(filename);
	//ws = wb->active_sheet();

	////统计总行数
	////最大行数减一，去掉列名
	//auto rows = ws.rows(false);
	//maxRow = rows.length() - 1;
	

	//maxRow = curRow = 1;
	//selColumns.clear();

	existingFile = false;;


	fileNames.clear();
	loadedWorkbook.clear();


}


void ExcelReader::loadXlsxFile(const std::string & pattern, const std::string & partOfSpeech, const std::string& path)
{
	std::regex re(pattern);
	for (auto& name : fileNames) {
		bool ret = std::regex_match(name, re);
		if (ret)
		{
			existingFile = true;	//设定已经加载文件

			std::string fullPath = path + name;
			
			for (auto& pair : loadedWorkbook) {
				if (pair.first == partOfSpeech) {
					xlnt::workbook workbook;
					
					workbook.load(fullPath);
					pair.second.push_back(workbook);
					return;
				}
			}
			//没有找到，说明还未添加该词类，则重新创建
			std::vector<xlnt::workbook> wbs;
			xlnt::workbook workbook;
			workbook.load(fullPath);
			wbs.push_back(workbook);

			loadedWorkbook.push_back(make_pair(partOfSpeech, wbs));

			break;

		}
	}
}

void ExcelReader::setPartOfSpeech(const std::string & str)
{
	curPartOfSpeech = str;

	changeWorkbook();

}

void ExcelReader::changeWorkbook()
{
	//默认选取std::vector<xlnt::workbook>中的第一个，如果有的话
	
	for (auto& pair : loadedWorkbook) {
		if (pair.first == curPartOfSpeech) {
			
			curWorkbook = &pair.second[0];
			curWorksheet = &curWorkbook->active_sheet();


			///////////////////////////////////////////////////////
			//重置
			selColumns.clear();
			curRow = 1;
			curWorkbookIndex = 0;

			//统计总行数
			//最大行数减一，去掉列名
			auto rows = curWorksheet->rows(false);
			maxRow = rows.length() - 1;


			return ;
		}
	}



}

bool ExcelReader::nextWorkbook()
{
	//


	return false;
}

bool ExcelReader::isExistingFile()
{
	return existingFile;
}

// 如果已达到最后一行，则返回false
bool ExcelReader::nextWord()
{



	/////////////////////////////////////////////////
	//TODO 切换到下一个工作簿






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




	//auto columns = ws.columns(false);
	//for (auto& column : columns) {
	//	std::string str = column[0].to_string();

	//	//使用xLnt读取xlsx文件，返回值均为utf-8编码
	//	//故str中实际存储的是utf-8编码的字符串

	//	if (columnName == str) {
	//		selColumns.push_back(make_pair(columnName, column));
	//		return;
	//	}

	//}

}

std::string ExcelReader::getCurCellValueInColumn(const std::string & columnName)
{
	//for (auto& pair : selColumns) {
	//	if (pair.first == columnName) {
	//		return pair.second[curRow].to_string();
	//	}
	//}

	return std::string("none");
}



