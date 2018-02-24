#include "DiceFace.h"
#include <string>
using namespace std;

DiceFace::DiceFace()
{
}

DiceFace::DiceFace(string c, string s) : Color(c), Symbol(s)
{
}

string DiceFace::getColor()
{
	return Color::toString();
}

string DiceFace::getSymbol()
{
	return Symbol::toString();
}

string  DiceFace::toString()
{
	return (getColor() + " " + getSymbol());
}
