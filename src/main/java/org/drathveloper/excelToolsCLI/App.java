package org.drathveloper.excelToolsCLI;

/**
 * Hello world!
 *
 */
public class App
{
    public static void main( String[] args )
    {
        Thread t = new Thread(new ExcelTools(args));
        t.start();
        synchronized (t){
            try {
                t.wait();
            } catch(InterruptedException ex){
                System.out.println("Thread error");
            }
        }
    }
}
