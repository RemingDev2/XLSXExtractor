#include "Randomizer.h"
#include <vector>
#include <algorithm>
#include <random>


std::vector<std::string> randomize(std::vector<std::string> List)
{
    std::mt19937 mersenTwister{std::random_device{}()};

    std::shuffle(List.begin(), List.end(), mersenTwister);

    return List;
}
