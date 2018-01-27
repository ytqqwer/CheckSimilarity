#pragma once

#include "xlnt/xlnt.hpp"

class ExcelReader
{
public:
	ExcelReader();
	~ExcelReader();

	void addXlsxFileName(const std::string& filename);

	void clear();
	
	void loadXlsxFile(const std::string& pattern, const std::string& partOfSpeech, const std::string& path);

	void setPartOfSpeech(const std::string&);

	bool changeWorkbook(unsigned int index = 0);
	bool nextWorkbook();

	bool nextWord();	// 如果已达到最后一行，则返回false	

	bool isExistingFile();

	void selectColumn(const std::string& columnName);
	std::string getCurCellValueInColumn(const std::string& columnName);

private:
	
	bool existingFile;

	std::string curPartOfSpeech;

	unsigned int curWorkbookIndex;
	
	//每次开始读取某一工作表前，重新设定行数
	unsigned int curRow;
	unsigned int maxRow;


	std::vector<std::pair<std::string, xlnt::cell_vector>> selColumns;	//已选择的列，需要指定列名
	
	std::vector<std::string> fileNames;
	std::vector<std::pair<std::string, std::vector<xlnt::workbook>>> loadedWorkbook;	 //已加载的工作簿
	
};




