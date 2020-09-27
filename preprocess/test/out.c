




int gmode1;














int calc(int x, int y)
{

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
