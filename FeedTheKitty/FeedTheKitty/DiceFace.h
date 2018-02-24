#pragma once
#ifndef DICEFACE_H
#define DICEFACE_H
#include "Color.h"
#include "Symbol.h"
#include <string>

using namespace std;

class DiceFace : public Color, public Symbol
{
public:
	DiceFace();
	DiceFace(string, string);
	string getColor();
	string getSymbol();
	string toString();
};
#endif