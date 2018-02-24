#include "Partition.h"
#include <assert.h>


Partition::Partition(GameTable *g, Partition *p)
{
	gtp = g;
	nextp = p;
	first = this;
}

Partition* Partition::getNextPartition()
{
		return nextp;
}

GameTable* Partition::getGameTable()
{
	if (gtp == nullptr)
		return NULL;
		return gtp;
}

void Partition::setNextPartition(Partition* p)
{
	nextp = p;
}

Partition::~Partition()
{
	if (first->getNextPartition() != NULL)
		delete first->getNextPartition();

		return;
}

void Partition::setGameTable(GameTable* g)
{
	gtp=g;
}