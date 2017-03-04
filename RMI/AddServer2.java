import java.rmi.*;
import java.net.*;
public class AddServer2
{
public static void main(String arg[])throws RemoteException
{
try
{
addmethod obj=new addmethod();
Naming.rebind("AddServer2",obj);
}
catch(Exception e)
{}
}
} 
