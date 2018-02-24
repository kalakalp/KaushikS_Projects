#include "TimingWheel.h"
int iteration = 0;
void TimingWheel::insert(int play_time, GameTable* g)
{

	if(m_slot[play_time]->getGameTable() == NULL && m_slot[play_time]->getNextPartition()==NULL)
	{
		m_slot[play_time]->setGameTable(g);
	}
	else
	{
		Partition* traverse = m_slot[play_time];
		while (traverse->getNextPartition() != NULL)
		{
			traverse = traverse->getNextPartition();
		}
		Partition *p = new Partition(g, NULL);

		traverse->setNextPartition(p);
	}
}

void TimingWheel::initialize()
{
	for (int i = 0; i < MAX_DELAY; i++)
	{
		if(m_slot[i]==NULL)
			m_slot[i]=new Partition(NULL, NULL);
		else
		{
			if (m_slot[i]->getGameTable() != NULL)
				m_slot[i]->setGameTable(NULL);
			if (m_slot[i]->getNextPartition() != NULL)
				m_slot[i]->setNextPartition(NULL);
		}
	}
}
void TimingWheel::clear_curr_slot()
{
	delete m_slot[m_current_slot];
	m_slot[m_current_slot]= NULL;
}
void TimingWheel::schedule(GameTable **g,int tables)
{
	initialize();
	iteration++;

	for (int i = 0; i < tables;i++)
	{
		int slot_index = g[i]->GetPlayerCount();
		insert(slot_index, g[i]);
	}
	
	for (int i = 0; i < MAX_DELAY; i++)
	{
		cout << endl << "##########################################"<<endl;
		cout << endl <<endl<<"Wheel slot :" << i+1 <<"X"<<iteration;
		if (m_slot[i]->getGameTable() != NULL)
		{
			m_slot[i]->getGameTable()->Play();
			Partition *p = m_slot[i];
			while (p->getNextPartition() != NULL)
			{
				p = p->getNextPartition();
				p->getGameTable()->Play();
			}
			m_current_slot = i;
			clear_curr_slot();
		}
	}
}
TimingWheel::TimingWheel()
{
	m_slot = new Partition*[MAX_DELAY];
	for (int i = 0; i < MAX_DELAY; i++)
	{
		m_slot[i] = new Partition(NULL, NULL);
	}
	m_current_slot = 0;

	cout << endl << "Wheel size : " << MAX_DELAY;
}