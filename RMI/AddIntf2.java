import java.rmi.*;
public interface AddIntf2 extends Remote
{
double add(double d1,double d2) throws RemoteException;
double mul(double d1,double d2) throws RemoteException;
double sub(double d1,double d2) throws RemoteException;
double div(double d1,double d2) throws RemoteException;
}