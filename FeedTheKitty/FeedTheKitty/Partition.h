#pragma once
#ifndef PARTITION_H
#define PARTITION_H
#include "GameTable.h"

class Partition {

	GameTable* gtp=NULL;
	Partition *nextp=NULL;
	Partition *first;

public:
	Partition(GameTable*, Partition*);
	~Partition();
	Partition* getNextPartition();
	void setNextPartition(Partition*);

	GameTable* getGameTable();
	void setGameTable(GameTable*);

};
#endif // !PARTITION_H
