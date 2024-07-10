#include <iostream>
#include "src/headers/XLSXWorker.h"


int main()
{
    std::cout << "Main Start" << std::endl;

    XLSXWorker doc("QCM_100");
    doc.extraction(true);

    std::cout << "Main Finished" << std::endl;

    return 0;
}
