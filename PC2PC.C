#include<stdio.h>


#define COM1       0
#define DATA_READY 0x100
#define TRUE       1
#define FALSE      0

#define SETTINGS ( 0x80 | 0x02 | 0x00 | 0x00)

int main(void)
{
   int out, status, DONE = FALSE;
   char ch = 's',file_name[20];
   char c;
   FILE *fp;

   clrscr();
   bioscom(0, SETTINGS, COM1);

   printf("... BIOSCOM [ESC] to exit ...\n");
   printf("\n Please Select your choice... \n 1.Transmit\n 2.Receive   :>");
   scanf("%d",&ch);

   bioscom(1,ch,COM1);

   if(ch == (bioscom(2,0,COM1)))
   {
    printf("\n Error Same choice at both the ends ");
    getch();
    exit(0);
   }

   switch(ch)
   {
    case 1:
	printf("\n Enter the file name to be transfer... :");
	scanf("%s",file_name);

	fp = fopen(file_name,"r");
	if(fp == NULL)
	{
	 printf("\n Error in File Operation");
	 exit(0);
	}
	else
	{
	  while(!feof(fp))
	  {
	   ch = fgetc(fp);

	   if(!feof(fp))
		bioscom(1,ch,COM1);

	   if(kbhit())
	   {
	     if(getche() == '\x1B')
	     break;
	   }
	  }
	}
	break;

    case 2:
	fp = fopen ("abc.txt","wb");

	if(fp == NULL)
	{
	 printf("\n Error in File Operation");
	 exit(0);
	}
     while (!DONE)
     {
	status = bioscom(3, 0, COM1);
	if (status & DATA_READY)
	 if ((out = bioscom(2, 0, COM1) & 0x7F) != 0)
	    putchar(out);
	 if (kbhit())
	 {
	    if ((ch = getch()) == '\x1B')
	       DONE = TRUE;
	//    bioscom(1, in, COM1);
	 }
     }
     break;


   default:
	printf("\n Wrong Choice is entered ......");
   }
   fclose(fp);
   getch();
   return (0);

}



