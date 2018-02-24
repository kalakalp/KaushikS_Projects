#pragma once
#ifndef SYMBOL_H
#define SYMBOL_H

#include<string>
using namespace std;

class Symbol
{
private:
	string m_faceSymbol;
public:
	Symbol();
	Symbol(string);
	virtual string toString();
};

#endif