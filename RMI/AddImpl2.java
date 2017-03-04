import java.rmi.*;
import java.rmi.server.*;
public class AddImpl2 extends UnicastRemoteObject implements AddIntf2
{
public AddImpl2()throws RemoteException
{}
public double add(double d1,double d2) throws RemoteException
{
return (d1+d2);
}
public double mul(double d1,double d2) throws RemoteException
{
return (d1*d2);
}
public double sub(double d1,double d2) throws RemoteException
{
return (d1-d2);
}
public double div(double d1,double d2) throws RemoteException
{
return (d1/d2);
}
}
