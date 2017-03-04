import java.rmi.*;
import java.net.*;
public class AddClient2
{
public static void main(String arg[])throws Exception
{
try
{
String addServerURL="rmi://"+arg[0]+"/AddServer2";
AddIntf2 oo=(AddIntf2)Naming.lookup(addServerURL);
double d1=Double.valueOf(arg[1]).doubleValue();
double d2=Double.valueOf(arg[2]).doubleValue();
System.out.print("Output \n Addition:");
System.out.println(oo.add(d1,d2));
System.out.print("Multiplication: ");
System.out.println(oo.mul(d1,d2));
System.out.print("Substraction: ");
System.out.println(oo.sub(d1,d2));
System.out.print("Division: ");
System.out.println(oo.div(d1,d2));
}
catch(Exception e)
{
System.out.println(e);
}
}
}