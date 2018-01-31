#include "stdafx.h"
#include "ExcelReader.h"
#include <codecvt>

#include <regex>	//正则表达式

#include <set>						

ExcelReader::ExcelReader()
{
	maxRow = curRow = curRowRangeBegin = curRowRangeEnd = 1;
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

bool ExcelReader::setPartOfSpeech(const std::string & str)
{
	curPartOfSpeech = str;

	//重新选择工作簿，并且跳过空表
	selColumns.clear();
	selRows.clear();
	curRow = 1;
	curRowRangeBegin = 1;
	curRowRangeEnd = 1;
	curWorkbookIndex = 0;

	return skipEmptyWorkbook();	
}

bool ExcelReader::skipEmptyWorkbook()
{
	for (auto& pair : loadedWorkbook) {
		if (pair.first == curPartOfSpeech) {
			unsigned int totalWorkbook = pair.second.size();
			
			if (curWorkbookIndex < totalWorkbook ) {
				
				xlnt::worksheet curWorksheet = pair.second[curWorkbookIndex].active_sheet();

				//最大行数减一，去掉列名
				auto rows = curWorksheet.rows(false);
				maxRow = rows.length() - 1;

				if (maxRow < curRow)
				{
					curWorkbookIndex++;
					return skipEmptyWorkbook();
				}
				else
				{
					//不需要再跳过，重新选择列
					selectColumn();
					//重新选择词组
					selectIsomorphicWordGroup();
					return true;
				}
			}
			else
				return false;
		}
	}

	return false;
}

//默认选取std::vector<xlnt::workbook>中的第一个，如果有的话
bool ExcelReader::changeWorkbook(unsigned int index)
{
	//重置
	selColumns.clear();
	selRows.clear();
	curRow = 1;
	curRowRangeEnd = 1;
	curWorkbookIndex = index;

	for (auto& pair : loadedWorkbook) {
		if (pair.first == curPartOfSpeech) {

			xlnt::worksheet curWorksheet = pair.second[index].active_sheet();

			//统计总行数，最大行数减一，去掉列名
			auto rows = curWorksheet.rows(false);
			maxRow = rows.length() - 1;

			if (maxRow < curRow)
			{
				return false; //说明当前工作簿只有一行列名，返回false
			}
			else
			{
				selectColumn();
				//重新选择词组
				selectIsomorphicWordGroup();
				return true;
			}
		}
	}

	return false;//什么也没找到
}

bool ExcelReader::nextWorkbook()
{
	for (auto& pair : loadedWorkbook) {
		if (pair.first == curPartOfSpeech) {
			unsigned int totalWorkbook = pair.second.size();

			if (curWorkbookIndex + 1 < totalWorkbook) {	//减1，防止索引越界
				curWorkbookIndex++;

				if (changeWorkbook(curWorkbookIndex)) {
					return true;
				}
				else {
					return nextWorkbook();
				}
			}
			else
				return false;//已没有下一个工作簿
		}
	}

	return false;//什么也没找到
}

//搜索词语同形的对应关系
void ExcelReader::selectIsomorphicWordGroup()
{
	bool isOver = false;
	curRowRangeBegin = curRowRangeEnd = curRow;	
	curWord = getCurCellValueInColumn(u8"gkb_词语");

	std::set<std::string> setOfIsomorphic;

	//获得词语的范围与同形数量
	while (!isOver) {
		std::string& word = getCurCellValueInColumn(u8"gkb_词语");
		if (curWord == word) 
		{
			setOfIsomorphic.insert(getCurCellValueInColumn(u8"gkb_同形"));
			curRow++;
			curRowRangeEnd++;
		}
		else 
		{
			isOver = true;
		}
	}
	
	//寻找同形所在行
	unsigned int row_of_isomorphic = 1;			//某一同形对应的词语所在行数，不一定是首个对应词语
	std::vector<unsigned int> matchedRows;	//匹配上的行
	for (auto& isomorphic : setOfIsomorphic) {
		for (curRow = curRowRangeBegin; curRow < curRowRangeEnd; curRow++)
		{			
			if (isomorphic == getCurCellValueInColumn(u8"gkb_同形")) {
				row_of_isomorphic = curRow;
				matchedRows.push_back(curRow);
			}
		}
		selRows.push_back(make_pair(row_of_isomorphic, matchedRows));
		matchedRows.clear();
	}
	
	numberOfIsomorphic = sizeOfSelectedRows();
	curIsomorphicIndex = 0;
}

// 如果已达到最后一行，则返回false
bool ExcelReader::nextWord()
{
	selRows.clear();

	if (curRowRangeEnd <= maxRow) {
		
		selectIsomorphicWordGroup();
		
		return true;
	}
	else
		return nextWorkbook();// 切换到下一个工作簿

}

bool ExcelReader::isExistingFile()
{
	return existingFile;
}

void ExcelReader::selectColumn()
{
	for (auto& columnName : columnNames) {
		for (auto& pair : loadedWorkbook) {
			if (pair.first == curPartOfSpeech) {

				xlnt::worksheet curWorksheet = pair.second[curWorkbookIndex].active_sheet();

				auto columns = curWorksheet.columns(false);
				for (auto& column : columns) {
					std::string str = column[0].to_string();

					//使用xLnt读取xlsx文件，返回值均为utf-8编码
					//故str中实际存储的是utf-8编码的字符串
					if (columnName == str) {
						selColumns.push_back(make_pair(columnName, column));
						break;
					}
				}
			}
		}
	}

}

std::string ExcelReader::getCurCellValueInColumn(const std::string & columnName)
{
	for (auto& pair : selColumns) {
		if (pair.first == columnName) {
			return pair.second[curRow].to_string();
		}
	}

	return std::string("none");
}

void ExcelReader::setIsomorphicColumnName(const std::string & columnName)
{
	isomorphicColumnName = columnName;
}

unsigned int ExcelReader::sizeOfSelectedRows()
{
	return selRows.size();
}

std::pair<unsigned int, std::vector<unsigned int>> ExcelReader::getRowByIndex(unsigned int index)
{
	unsigned int i = 0;
	for (auto& pair : selRows) {
		if (i == index) {
			return pair;
		}
		else
			i++;
	}

	return std::pair<unsigned int, std::vector<unsigned int>>();
}

std::string ExcelReader::getValueInColumnByRow(unsigned int row, const std::string & columnName)
{
	for (auto& pair : selColumns) {
		if (pair.first == columnName) {
			return pair.second[row].to_string();
		}
	}

	return std::string("none");
}

void ExcelReader::setColumnNames(const std::vector<std::string>& names)
{
	columnNames = names;
}