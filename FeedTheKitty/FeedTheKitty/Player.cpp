#include "Player.h"
int Player::ms_Playercount = 0;

Player::Player()
{
	ms_Playercount++;
}

Player::Player(int id,int micecount)
{
	setId(id);
	setMiceCount(micecount);
}

void Player::setId(int id)
{
		m_Pid = id;
}

void Player::setMiceCount(int count)
{
		m_MiceCount = count;
}

int Player::getId()
{
	return m_Pid;
}

int Player::getMiceCount()
{
	return m_MiceCount;
}