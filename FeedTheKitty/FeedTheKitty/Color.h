#pragma once
#ifndef COLOR_H
#define COLOR_H

#include<string>
using namespace std;

class Color
{
private:
	string m_faceColor;
public:
	Color();
	Color(string);
	virtual string toString();
};
#endif