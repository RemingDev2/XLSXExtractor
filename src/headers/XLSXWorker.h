#ifndef KUESTO_XLSXWORKER_H
    #define KUESTO_XLSXWORKER_H

#include <string>
#include <vector>
#include <OpenXLSX.hpp>


class XLSXWorker {
public:
    XLSXWorker(const std::string& excelFileName);

    void extraction(bool is_onRow);

private:
    OpenXLSX::XLDocument doc;
    OpenXLSX::XLWorksheet currentWorkSheet;

    std::vector<std::string> questionsList;
    std::vector<std::string> answersList;

    void rowAnswers(const std::string& questionRowNumber);
    void columnAnswers(int questionRowNumber);
};


#endif
