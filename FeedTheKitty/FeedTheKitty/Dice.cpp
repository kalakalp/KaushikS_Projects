#include "Dice.h"
#include <fstream>
#include <string>
#include <iostream>
#include <ctime>

using namespace std;

Dice::Dice()
{
	m_Count[0] = 0;
	m_Count[1] = 0;
	m_Count[2] = 0;
}

Dice::Dice(string filename)
{
	m_colorMap = new map<int, string>;
	m_symbolMap = new map<int, string>;

	initialize(filename);
	int facecount = m_Count[0];
	m_faces = new DiceFace*[facecount];

	for (int i = 0; i<facecount; i++)
		m_faces[i] = createDiceface();
}

Dice::~Dice()
{
	int facecount = m_Count[0];
	for (int i = 0; i < facecount; i++)
		delete m_faces[i];
	delete[] m_faces;
}

DiceFace* Dice::createDiceface()
{
	int random = 0;
	string color, symbol;


	random = (rand()*time(NULL)) % 100;

	for (map<int, string>::iterator it = m_colorMap->begin(); it != m_colorMap->end(); ++it)
	{
		if (random < ((*it).first))
		{
			color = (*it).second;
			break;
		}
	}

	random = (rand()*time(NULL)) % 100;

	for (map<int, string>::iterator it = m_symbolMap->begin(); it != m_symbolMap->end(); ++it)
	{
		if (random < ((*it).first))
		{
			symbol = (*it).second;
			break;
		}
	}
	return (new DiceFace(color, symbol));
}
void Dice::initialize(string file)
{
	ifstream in("..\\Resources\\"+file);
	char m_CountChar;
	
	if (in.is_open())
	{
		int i = 0, prob = 0;
		string l_colorOrsymbol, l_prob;

		for (; i < 3; i++)
		{

			in >> m_CountChar;
			m_Count[i] = m_CountChar - '0';
		}
		int count = m_Count[1];

		for (i = 0; (i < count && !in.eof()); i++)
		{
			in >> l_colorOrsymbol;
			in >> l_prob;
			prob += stoi(l_prob);
			m_colorMap->insert(make_pair(prob, l_colorOrsymbol));
		}

		prob = 0;
		count = m_Count[2];

		for (i = 0; (i < count && !in.eof()); i++)
		{
			in >> l_colorOrsymbol;
			in >> l_prob;
			prob += stoi(l_prob);
			m_symbolMap->insert(make_pair(prob, l_colorOrsymbol));
		}
	}
	else
		cout << endl << "Error: Unable to open file, try a different file";
}

void Dice::display()
{
	int facecount = m_Count[0];
	for (int i = 0; i < facecount; i++)
		cout << m_faces[i]->toString() << "\t";
}

string Dice::RollDice()
{
	return m_faces[((rand()*time(NULL)) % m_Count[0])]->toString();
}