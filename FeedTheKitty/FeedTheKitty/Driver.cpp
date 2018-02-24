#include<iostream>
#include<ctime>
#include "GameTable.h"
#include "TimingWheel.h"
using namespace std;

int main()
{
	int tables;
	cout << "Enter the number of tables in the arcade" << endl;
	cin >> tables;
	
	int players = 0;

	TimingWheel TM;
	int count = 3;
	GameTable **g;

	while (count-- > 0)
	{
		
		g = new GameTable*[tables];
		for (int i = 0; i < tables; i++)
		{
			players = (rand()*time(NULL)) % 7;
			if (players < 2)
			{
				i--;
				continue;
			}
			g[i] = new GameTable(players, 2, 20, "input.txt", to_string(i+1));
		}

		TM.schedule(g, tables);

		for (int i = 0; i < tables; i++)
		{
			delete g[i];
		}
		delete[] g;
		g = NULL;
	}

	getchar();
	return 0;
}