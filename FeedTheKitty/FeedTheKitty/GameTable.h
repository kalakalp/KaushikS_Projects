#pragma once
#ifndef GAMETABLE_H
#define GAMETABLE_H
#include<iostream>
#include "Game.h"

class GameTable : public Game
{
	string m_name;
public:
	GameTable(int, int, int, string, string);
	~GameTable();

	void Play();
	void setName(string);
	string getName();
};
#endif
