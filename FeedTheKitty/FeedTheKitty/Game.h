#pragma once
#ifndef GAME_H
#define GAME_H
#include<string>
#include<vector>
#include "Dice.h"
#include "Player.h"
using namespace std;

class Game
{
	int m_Players;
	int m_DiceCount;
	int m_MiceCount;
	int m_Bowl;
	string m_name;
	Player *m_PlayerArray;
	vector<Dice*> m_Dice;

public:
	Game(string);
	~Game();
	void SetPlayers(int);
	int GetPlayerCount();
	void SetDiceCount(int);
	void SetMiceCount(int);
	void CreateDice(string);
	virtual void Play();
	void RollDice();
	int Action(int);
	void Results();
	void Display(int, int);
	void Initialize();
};
#endif