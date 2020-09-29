#include <stdio.h>
#include "option.h"

#if (defined(X_SYS) || defined(Y_SYS))
#include "xxx.h"
#endif
#define mode1
#if defined(mode1)
int gmode1;
#elif defined(mode2)
int gmode2;
#else
int gmode0;
#endif
#ifdef DEBUG
int gdebug;
#endif // DEBUG
#ifndef option1
int gone;
#endif // !1


/*comment */
int calc(int x, int y)
{
	/*comment 2*/
	return x + y;
}
int mul(int x, int y)
{
	return x * y;
}

int main()
{
	int x;
	int i;
	for (i = 0;i<10;i++)
	{
		if (i%2 == 0) {
			x = calc(10,11);
		} else {
			x = mul(i,2);
		}
	}
	return 0;
}