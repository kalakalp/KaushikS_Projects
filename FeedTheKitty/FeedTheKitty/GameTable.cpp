#include "GameTable.h"

GameTable::GameTable(int players = 4, int dice = 2, int mice = 20, string file = "input.txt", string nameId=0):Game(nameId)
{
		SetPlayers(players);
		SetDiceCount(dice);
		SetMiceCount(mice);
		CreateDice(file);
		setName("Game Table "+nameId);
		Initialize();
}

GameTable::~GameTable()
{
}
void GameTable::Play()
{
	cout <<endl<< "Playing game on " << m_name;
	cout << endl << "\t***************"<<endl ;

	Game::Play();
	Results();
}
void GameTable::setName(string s)
{
	m_name = s;
}
string GameTable::getName()
{
	return m_name;
}
