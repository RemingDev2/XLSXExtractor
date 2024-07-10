#include "../headers/XLSXWorker.h"
#include "../../external/Randomizer/Randomizer.h"
#include <string>
#include <SQLiteCpp/SQLiteCpp.h>


// constructor
XLSXWorker::XLSXWorker(const std::string& excelFileName): doc(excelFileName + ".xlsx")
{
    // open the first sheet ||| worksheet() = open ||| worksheetName() = list of name of sheets
    currentWorkSheet = doc.workbook().worksheet(doc.workbook().worksheetNames()[0]);
}


// get answers (if answers are in the same row as the question or in column)
void XLSXWorker::rowAnswers(const std::string& questionRowNumber)
{
    std::string upperCase = "CDEFGHIJKL";

    int columnIndex(0);

    do {
        OpenXLSX::XLCell currentCellValue = currentWorkSheet.cell(upperCase[columnIndex] + questionRowNumber);

        switch (currentCellValue.value().type())
        {
            case OpenXLSX::XLValueType::String:
                answersList.push_back(currentCellValue.value());
                break;
            case OpenXLSX::XLValueType::Integer:
                answersList.push_back(std::to_string(currentCellValue.value().get<int>()));
                break;
            case OpenXLSX::XLValueType::Float:
                answersList.push_back(std::to_string(currentCellValue.value().get<float>()));
                break;
        }

        ++columnIndex;
    } while (currentWorkSheet.cell(upperCase[columnIndex] + questionRowNumber).value().type() != OpenXLSX::XLValueType::Empty
             && columnIndex <= currentWorkSheet.columnCount());
}

void XLSXWorker::columnAnswers(int questionRowNumber)
{
    int rowIndex(questionRowNumber);

    do {
        OpenXLSX::XLCell currentCell = currentWorkSheet.cell("C" + std::to_string(rowIndex));

        switch (currentCell.value().type())
        {
            case OpenXLSX::XLValueType::String:
                answersList.push_back(currentCell.value());
                break;
            case OpenXLSX::XLValueType::Integer:
                answersList.push_back(std::to_string(currentCell.value().get<int>()));
                break;
            case OpenXLSX::XLValueType::Float:
                answersList.push_back(std::to_string(currentCell.value().get<float>()));
                break;
        }

        ++rowIndex;

    } while (currentWorkSheet.cell("B"+std::to_string(rowIndex)).value().type() == OpenXLSX::XLValueType::Empty
             && currentWorkSheet.cell("C"+std::to_string(rowIndex)).value().type() != OpenXLSX::XLValueType::Empty);
}


// extract data
void XLSXWorker::extraction(bool is_onRow)
{
    SQLite::Database db("database.db", SQLite::OPEN_READWRITE);

    db.exec("DELETE FROM datas; DELETE FROM sqlite_sequence; VACUUM;");

    for (int i(1); i <= currentWorkSheet.rowCount(); ++i)
    {
        OpenXLSX::XLCell currentCell = currentWorkSheet.cell("B"+std::to_string(i));

        if (currentCell.value().type() != OpenXLSX::XLValueType::Empty)
        {
            SQLite::Transaction transac(db);

            SQLite::Statement query(db, "INSERT INTO datas VALUES "
                                        "(null, ?, ?, null, null, null, null, null, null, null, null, null, null);");

            ///
            std::string ctrBind = currentWorkSheet.cell("A" + std::to_string(i)).value();
            query.bind(1, ctrBind);

            ///
            std::string qtnBind = currentCell.value();
            query.bind(2, qtnBind);

            query.exec(); query.reset(); query.clearBindings();

            transac.commit();

            if (is_onRow) {rowAnswers(std::to_string(i));} else {columnAnswers(i);}

            SQLite::Transaction asrTransac(db);

            int incr(1);
            for (std::string& asrCurrent : randomize(answersList))
            {
                SQLite::Statement asrQuery(db, "UPDATE datas "
                                               "SET asr" + std::to_string(incr) + " = ? "
                                               "WHERE ctr = ? AND qtn = ?;");

                asrQuery.bind(1, asrCurrent); asrQuery.bind(2, ctrBind); asrQuery.bind(3, qtnBind);
                asrQuery.exec();
                incr++;
            }
            asrTransac.commit();
            answersList.clear();
        }
    }
}
