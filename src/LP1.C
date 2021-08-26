#define bal 99999.
#define payment -70554.045610249
#define rate .705547511577606
	double dbal = bal;
	long double ldbal = bal;
	int i;
	for(i=1;i<=66;++i)
	{
		dbal = dbal + payment + dbal * rate;
		ldbal = ldbal + payment + ldbal * rate;
		printf("dbal=%lf ldbal=%lf\n",dbal,ldbal);
	}
