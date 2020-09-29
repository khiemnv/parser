#if !defined(opt_inc)

#define lvl2	/*lvl2*/
/*#define lvl3*/	/*lvl3*/

#if defined(lvl5)
/*#define xxx*/	/*xxx*/
#endif
#if (!defined(X_SYS) && !defined(Y_SYS))
  #define JUG_X	/*xxx*/
#endif

#define opt_inc

#define option1

#if defined(option1)
int opt1;
#else
int opt0;
#endif
#endif

