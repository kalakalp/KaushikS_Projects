#pragma once
#ifndef DICE_H
#define DICE_H
#include "DiceFace.h"
#include <map>

class Dice
{
private:

	map<int, string> *m_colorMap;
	map<int, string> *m_symbolMap;
	int m_Count[3];
	DiceFace **m_faces;

public:
	static int m_ran;
	Dice();
	~Dice();
	Dice(string);
	void initialize(string);
	DiceFace* createDiceface();
	void display();
	string RollDice();
};
#endif