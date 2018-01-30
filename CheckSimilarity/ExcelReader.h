#pragma once

#include "xlnt/xlnt.hpp"

class ExcelReader
{
public:
	ExcelReader();
	~ExcelReader();

	void addXlsxFileName(const std::string& filename);
	void loadXlsxFile(const std::string& pattern, const std::string& partOfSpeech, const std::string& path);
	void clear();
	
	bool setPartOfSpeech(const std::string&);

	void selectIsomorphicWordGroup();
	bool nextWord();

	void setColumnNames(const std::vector<std::string>& columnNames);
	
	bool isExistingFile();

	void setIsomorphicColumnName(const std::string& columnName);
	
	std::pair<unsigned int, std::vector<unsigned int>> getRowByIndex(unsigned int);

	std::string getValueInColumnByRow(unsigned int row,const std::string& columnName);
	
public:	
	unsigned int numberOfIsomorphic;
	unsigned int curIsomorphicIndex;
	
private:

	bool skipEmptyWorkbook();
	bool changeWorkbook(unsigned int index = 0);
	bool nextWorkbook();

	void selectColumn();

	std::string getCurCellValueInColumn(const std::string& columnName);

	unsigned int sizeOfSelectedRows();

private:	
	bool existingFile;
	unsigned int curWorkbookIndex;
	
	std::string curPartOfSpeech;

	std::vector<std::string> fileNames;
	std::vector<std::pair<std::string, std::vector<xlnt::workbook>>> loadedWorkbook;	 //已加载的工作簿

	std::string isomorphicColumnName;

private:
	std::string curWord;	//当前词语

	
	unsigned int curRow;		
	unsigned int curRowRangeBegin;//当前词组范围
	unsigned int curRowRangeEnd;
	
	unsigned int maxRow;

	std::vector<std::string> columnNames;
	std::vector<std::pair<std::string, xlnt::cell_vector>> selColumns;	//已选择的列，需指定列名
	
	std::vector<std::pair<unsigned int , std::vector<unsigned int>>> selRows;	//同形和行数的对应
		
};
