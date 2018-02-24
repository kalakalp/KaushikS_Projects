#include<iostream>
#include <fstream>
#include "Game.h"
#include<Windows.h>
#include <ctime>
#include "GameFlags.h"
using namespace std;

vector<string> DiceOutcome;
int CurrentPlayer = 0;
bool Direction = false;
bool PreviousDirection = false;


Game::Game(string name)
{
	m_name = name;

	m_Players = m_DiceCount = m_MiceCount = 0;
	vector<Dice> *m_Dice = NULL;
}

Game::~Game()
{
	try {
		delete[] m_PlayerArray;
	}
	catch(...)
	{ }
}
void Game::SetPlayers(int Players)
{
	m_Players = Players;
	m_PlayerArray = new Player[m_Players];
}
int Game::GetPlayerCount()
{
	return m_Players;
}
void Game::SetDiceCount(int Dice)
{
	m_DiceCount = Dice;
	m_Dice.reserve(m_DiceCount);
}
void Game::SetMiceCount(int Mice)
{
	m_MiceCount = Mice;
}
void Game::CreateDice(string file)
{
	for (int i = 0; i < m_DiceCount; i++)
	{
		m_Dice.push_back(new Dice(file));
	}
}

void Game::Initialize()
{
	int MicePerPlayer = m_MiceCount / m_Players;
	for (int i = 0; i < m_Players; i++)
	{
		m_PlayerArray[i].setId(i);
		m_PlayerArray[i].setMiceCount(MicePerPlayer);
	}
	m_Bowl = m_MiceCount%m_Players;
}
void Game::Play() {
	
	int loopcount = 0;
	
	while (loopcount++<30)	
	{
		int temp = CurrentPlayer;

		RollDice(); //Rolling all the dice
		CurrentPlayer = Action(CurrentPlayer); //Update the mice count of each player
		Display(temp, CurrentPlayer); //Print the status

		DiceOutcome.clear();
	}
}

int Game::Action(int cp)
{

	Flags flags;

	for (vector<string>::iterator it = DiceOutcome.begin(); it != DiceOutcome.end(); ++it)
	{

		flags.DiceResult = *it;
		if (it == DiceOutcome.begin())
			flags.color = flags.DiceResult[0];

		PreviousDirection = Direction;
		if (flags.DiceResult[0] != flags.color)
			if (Direction == true)
				Direction = false;
			else
				Direction = true;

		if ((flags.DiceResult[2] == 'S'&& flags.DiceResult[3] == 'c') && (flags.TotheBowl == 0) && (flags.Tothenextplayer == 0) && (flags.Takefrombowl == 0))
			flags.Donothing = true;
		else
			flags.Donothing = false;

		if (flags.DiceResult[2] == 'B')
			flags.TotheBowl++;

		if (flags.DiceResult[2] == 'A')
			flags.Tothenextplayer++;

		if (flags.DiceResult[2] == 'M')
			flags.Takefrombowl++;
	}


	int nextplayer = 0;
	if (PreviousDirection)
	{
		nextplayer = ((cp - 1) % m_Players);
		if (nextplayer < 0)
			nextplayer = m_Players + nextplayer;
	}
	else
		nextplayer = ((cp + 1) % m_Players);

	if (flags.Donothing)
	{
		return nextplayer;
	}

	if (flags.TotheBowl > 0)
	{
		if (m_PlayerArray[cp].getMiceCount() >= flags.TotheBowl)
		{
			m_PlayerArray[cp].setMiceCount(m_PlayerArray[cp].getMiceCount()- flags.TotheBowl);
			m_Bowl += flags.TotheBowl;
		}
	}

	if (flags.Takefrombowl > 0)
	{
		if (m_Bowl >= flags.Takefrombowl)
		{
			m_Bowl -= flags.Takefrombowl;
			m_PlayerArray[cp].setMiceCount(m_PlayerArray[cp].getMiceCount() + flags.Takefrombowl);
		}
	}
	if (flags.Tothenextplayer> 0)
	{
		if (m_PlayerArray[cp].getMiceCount() >= flags.Tothenextplayer)
		{

			m_PlayerArray[cp].setMiceCount(m_PlayerArray[cp].getMiceCount() - flags.Tothenextplayer);
			m_PlayerArray[nextplayer].setMiceCount(m_PlayerArray[nextplayer].getMiceCount() + flags.Tothenextplayer);
		}
	}


	nextplayer = 0;
	if (Direction)
	{
		nextplayer = ((cp - 1) % m_Players);
		if (nextplayer < 0)
			nextplayer = m_Players + nextplayer;
	}
	else
		nextplayer = ((cp + 1) % m_Players);

	return nextplayer;
}

void Game::Display(int cp, int np)
{

	ofstream fout;
	fout.open("Output\\Table "+m_name+".txt",fstream::app);

	fout << endl;
	for (vector<string>::iterator it = DiceOutcome.begin(); it != DiceOutcome.end(); ++it)
	{
		fout << *it << "\t";
	}
	fout << endl << "Current Player: P" << cp + 1 << endl;
	fout << "Next Player: P" << np + 1 << endl;
	fout << "Mice: B-" << m_Bowl;

	for (int i = 0; i < m_Players; i++)
		fout << ", P" << i + 1 << "-" << m_PlayerArray[i].getMiceCount();

	fout.close();
}
void Game::Results()
{
	ofstream fout;
	fout.open("Output\\Table " + m_name + ".txt", fstream::app);
	fout << endl << endl << "\t***************" << endl << endl;

	int big = 0;
	for (int i = 0; i < m_Players; i++)
	{
		if (m_PlayerArray[i].getMiceCount() >= big)
			big = m_PlayerArray[i].getMiceCount();
	}

	cout << "\tWinner(s): ";
	fout << "\tWinner(s): ";

	for (int i = 0; i < m_Players; i++)
	{
		if (m_PlayerArray[i].getMiceCount() == big)
		{
			cout << "P" << i + 1 << "\t";
			fout << "P" << i + 1 << "\t";
		}
	}
	cout << endl << "\t***************"<<endl ;

	fout << endl ;

}

void Game::RollDice()
{
	for (int i = 0; i < m_DiceCount; i++)
	{
		DiceOutcome.push_back(m_Dice[i]->RollDice());
	}
}
