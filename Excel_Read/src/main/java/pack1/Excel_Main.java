package pack1;

import java.io.IOException;
public class Excel_Main
{
public static void main(String[] args) throws IOException
{
String s=Excel_new.getStringData(0, 1);  //changed
System.out.println(s);
String n=Excel_new.getIntegerData(0,0);
System.out.println(n);
}
}
