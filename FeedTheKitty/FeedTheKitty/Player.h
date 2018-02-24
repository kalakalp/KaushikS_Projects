#pragma once
#ifndef PLAYER_H
#define PLAYER_H

class Player
{
	int m_Pid;
	int m_MiceCount;
public:
	static int ms_Playercount;

	Player();
	Player(int,int);
	void setId(int);
	void setMiceCount(int);
	int getId();
	int getMiceCount();
};
#endif