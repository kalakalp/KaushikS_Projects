#pragma once
#ifndef GAMEFLAGS_H
#define GAMEFLAGS_H
#include<string>
using namespace std;

struct Flags
{
	bool Donothing = false;
	int TotheBowl = 0;
	int Tothenextplayer = 0;
	int Takefrombowl = 0;
	char color;
	string DiceResult = "";
};
#endif // !GAMEFLAGS_H
