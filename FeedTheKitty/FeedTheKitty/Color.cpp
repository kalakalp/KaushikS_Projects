#include "Color.h"

Color::Color()
{
	m_faceColor = "";
}

Color::Color(string p_color)
{
	m_faceColor = p_color;
}

string Color::toString()
{
	return m_faceColor;
}