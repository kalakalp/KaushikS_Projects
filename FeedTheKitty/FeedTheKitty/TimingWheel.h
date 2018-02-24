#pragma once
#ifndef TIMINGWHEEL_H
#define TIMINGWHEEL_H

#include "Partition.h"
#include "GameTable.h"
#define MAX_DELAY 10

class TimingWheel {

	Partition **m_slot;// [MAX_DELAY + 1];
	int m_current_slot;

public:
	TimingWheel();
	void insert(int play_time, GameTable* p1);
	void schedule(GameTable**,int);
	void clear_curr_slot();
	void initialize();
};

#endif // !TIMINGWHEEL_H
