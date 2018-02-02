#include "stdafx.h"
#include "ExcelReader.h"
#include <codecvt>

#include <regex>	//正则表达式

#include <set>						

ExcelReader::ExcelReader()
{
	highestRow = curRow = curRowRangeBegin = curRowRangeEnd = 1;
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

void ExcelReader::reset()
{
	selColumns.clear();
	selRows.clear();
	curRow = 1;
	curRowRangeBegin = 1;
	curRowRangeEnd = 1;
	curWorkbookIndex = 0;
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
	reset();

	return skipEmptyWorkbook();
}

bool ExcelReader::skipEmptyWorkbook()
{
	for (auto& pair : loadedWorkbook) {
		if (pair.first == curPartOfSpeech) {
			unsigned int totalWorkbook = pair.second.size();

			if (curWorkbookIndex < totalWorkbook) {
				xlnt::worksheet& curWorksheet = pair.second[curWorkbookIndex].active_sheet();

				//highestRow = curWorksheet.highest_row();

				auto& rows = curWorksheet.rows();
				int i = 0;
				for (auto& word : rows) {					
					if (word[0].to_string() != u8"") {
						i++;
					}
				}
				highestRow = i;



				if (highestRow > curRow)
				{
					selectColumn(curWorkbookIndex);//不需要再跳过，重新选择列					
					selectNextIsomorphicWordGroup(curRow);//重新选择词组
					return true;
				}
				else
				{
					curWorkbookIndex++;
					return skipEmptyWorkbook();
				}
			}
			else
				return false;
		}
	}
	return false;
}

bool ExcelReader::changeWorkbook(unsigned int index)
{
	reset();//重置
	curWorkbookIndex = index;

	for (auto& pair : loadedWorkbook) {
		if (pair.first == curPartOfSpeech) {

			xlnt::worksheet& curWorksheet = pair.second[index].active_sheet();

			//highestRow = curWorksheet.highest_row();

			auto& rows = curWorksheet.rows();
			int i = 0;
			for (auto& word : rows) {
				if (word[0].to_string() != u8"") {
					i++;
				}
			}
			highestRow = i;


			if (highestRow > curRow)
			{
				selectColumn(index);
				return true;
			}
			else
			{
				return false; //说明当前工作簿只有一行列名，返回false
			}
		}
	}
	return false;//什么也没找到
}

bool ExcelReader::prevWorkbook()
{
	for (auto& pair : loadedWorkbook) {
		if (pair.first == curPartOfSpeech) {

			if (0 < curWorkbookIndex) {
				curWorkbookIndex--;

				if (changeWorkbook(curWorkbookIndex)) {
					//重新选择词组
					selectPreviousIsomorphicWordGroup(highestRow - 1);
					return true;
				}
				else {
					return prevWorkbook();
				}
			}
			else
				return false;
		}
	}

	return false;//什么也没找到
}

bool ExcelReader::nextWorkbook()
{
	for (auto& pair : loadedWorkbook) {
		if (pair.first == curPartOfSpeech) {
			unsigned int totalWorkbook = pair.second.size();

			if (curWorkbookIndex < totalWorkbook - 1) {	//防止索引越界
				curWorkbookIndex++;

				if (changeWorkbook(curWorkbookIndex)) {
					//重新选择词组
					selectNextIsomorphicWordGroup(1);
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

void ExcelReader::selectPreviousIsomorphicWordGroup(unsigned int row)
{
	selRows.clear();

	curRowRangeBegin = row;
	curRowRangeEnd = row;

	//获得词语的范围与同形数量
	std::set<std::string> setOfIsomorphic;
	curWord = getCurCellValueInColumn(row, u8"gkb_词语");
	bool isOver = false;
	while (!isOver)
	{
		std::string& word = getCurCellValueInColumn(row, u8"gkb_词语");
		if (curWord == word && word != u8"")
		{
			setOfIsomorphic.insert(getCurCellValueInColumn(row, u8"gkb_同形"));
			row--;
		}
		else
		{
			isOver = true;
		}
	}

	curRowRangeBegin = row + 1;

	//寻找同形所在行
	unsigned int row_of_isomorphic = highestRow;			//某一同形对应的词语所在行数，不一定是首个对应词语
	std::vector<unsigned int> matchedRows;	//匹配上的行
	for (auto& isomorphic : setOfIsomorphic) {
		for (row = curRowRangeBegin; row <= curRowRangeEnd; row++)
		{
			if (isomorphic == getCurCellValueInColumn(row, u8"gkb_同形")) {
				row_of_isomorphic = row;
				matchedRows.push_back(row);
			}
		}
		selRows.push_back(make_pair(row_of_isomorphic, matchedRows));
		matchedRows.clear();
	}

	numberOfIsomorphic = sizeOfSelectedRows();
	curIsomorphicIndex = 0;
	curRow = row;
}

//搜索词语同形的对应关系
void ExcelReader::selectNextIsomorphicWordGroup(unsigned int row)
{
	selRows.clear();

	curRowRangeBegin = row;
	curRowRangeEnd = row;

	std::set<std::string> setOfIsomorphic;

	//获得词语的范围与同形数量
	curWord = getCurCellValueInColumn(row, u8"gkb_词语");
	bool isOver = false;
	while (!isOver) {
		std::string& word = getCurCellValueInColumn(row, u8"gkb_词语");
		if (curWord == word && word != u8"")
		{
			setOfIsomorphic.insert(getCurCellValueInColumn(row, u8"gkb_同形"));
			row++;
		}
		else
		{
			isOver = true;
		}
	}
	curRowRangeEnd = row - 1;

	//寻找同形所在行
	unsigned int row_of_isomorphic = 1;			//某一同形对应的词语所在行数，不一定是首个对应词语
	std::vector<unsigned int> matchedRows;	//匹配上的行
	for (auto& isomorphic : setOfIsomorphic) {
		for (row = curRowRangeBegin; row <= curRowRangeEnd; row++)
		{
			if (isomorphic == getCurCellValueInColumn(row, u8"gkb_同形")) {
				row_of_isomorphic = row;
				matchedRows.push_back(row);
			}
		}
		selRows.push_back(make_pair(row_of_isomorphic, matchedRows));
		matchedRows.clear();
	}

	numberOfIsomorphic = sizeOfSelectedRows();
	curIsomorphicIndex = 0;
	curRow = row;
}

bool ExcelReader::prevWord()
{
	if (0 < curRowRangeBegin - 1) {
		selectPreviousIsomorphicWordGroup(curRowRangeBegin - 1);
		return true;
	}
	else
		return prevWorkbook();// 切换到上一个工作簿
}

bool ExcelReader::nextWord()
{
	if (curRowRangeEnd + 1 < highestRow) {
		selectNextIsomorphicWordGroup(curRowRangeEnd + 1);
		return true;
	}
	else
		return nextWorkbook();// 切换到下一个工作簿
}

bool ExcelReader::isExistingFile()
{
	return existingFile;
}

void ExcelReader::selectColumn(unsigned int workbookIndex)
{
	for (auto& columnName : columnNames) {
		for (auto& pair : loadedWorkbook) {
			if (pair.first == curPartOfSpeech) {

				xlnt::worksheet& curWorksheet = pair.second[workbookIndex].active_sheet();

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

std::string ExcelReader::getCurCellValueInColumn(unsigned int row, const std::string & columnName)
{
	for (auto& pair : selColumns) {
		if (pair.first == columnName) {
			return pair.second[row].to_string();
		}
	}

	return std::string("none");
}

void ExcelReader::setIsomorphicColumnName(const std::string & columnName)
{
	isomorphicColumnName = columnName;
}

void ExcelReader::setWordColumnName(const std::string & columnName)
{
	wordColumnName = columnName;
}

unsigned int ExcelReader::sizeOfSelectedRows()
{
	return selRows.size();
}

std::pair<unsigned int, std::vector<unsigned int>> ExcelReader::getRowsByIndex(unsigned int index)
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

bool ExcelReader::findWord(const std::string & word)
{
	unsigned int preCurRow = curRow;
	unsigned int preWorkbookIndex = curWorkbookIndex;

	for (auto& pair : loadedWorkbook) {
		if (pair.first == curPartOfSpeech) {
			unsigned int totalWorkbook = pair.second.size();

			for (curWorkbookIndex = 0; curWorkbookIndex < totalWorkbook; curWorkbookIndex++)
			{
				if (curWorkbookIndex < totalWorkbook) {
					changeWorkbook(curWorkbookIndex);

					for (auto& column : selColumns)
					{
						if (column.first == wordColumnName)
						{
							for (auto& str : column.second)
							{
								if (str.to_string() == word)
								{
									curRow--;//列中数据包含了列名，去掉因为列名多加的1

									selectNextIsomorphicWordGroup(curRow);
									return true;
								}
								else {
									curRow++;
								}
							}
							break;//不需要再遍历其他列
						}
					}

				}
				else
					return false;

			}

		}
	}

	curWorkbookIndex = preWorkbookIndex;
	curRow = preCurRow;

	changeWorkbook(curWorkbookIndex);
	selectNextIsomorphicWordGroup(curRow);

	return false;
}

void ExcelReader::setColumnNames(const std::vector<std::string>& names)
{
	columnNames = names;
}