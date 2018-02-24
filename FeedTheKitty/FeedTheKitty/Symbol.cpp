#include "Symbol.h"

Symbol::Symbol()
{
	m_faceSymbol = "";
}

Symbol::Symbol(string p_Symbol)
{
	m_faceSymbol = p_Symbol;
}

string Symbol::toString()
{
	return m_faceSymbol;
}