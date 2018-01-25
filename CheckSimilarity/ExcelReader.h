#pragma once

#include "xlnt/xlnt.hpp"

class ExcelReader
{
public:
	ExcelReader();
	~ExcelReader();

	void loadXlsxFile(const std::string& filename);
	bool isOpenFile();
	bool nextWord();	// 如果已达到最后一行，则返回false	
	void selectColumn(const std::string& columnName);
	std::string getCurCellValueInColumn(const std::string& columnName);

private:
	bool isOpen;

	xlnt::workbook* wb;
	xlnt::worksheet ws;

	unsigned int curRow;
	unsigned int maxRow;

	std::vector<std::pair<std::string, xlnt::cell_vector>> selColumns;	//已选择的列，需要指定列名
	
};




