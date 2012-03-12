package ua.org.pycb6a.pcplreport;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;

public class MainFrame {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub

		EventQueue.invokeLater (new Runnable()
		{
			public void run()
			{
		ReportFrame mainFrame = new ReportFrame();
		mainFrame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		mainFrame.setLocationRelativeTo(null);
		mainFrame.setResizable(false);

		 try
		 {
		        UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
		        SwingUtilities.updateComponentTreeUI(mainFrame);
		     		      
		 }
		 catch (Exception e)
		 {
			 e.printStackTrace();
		 }
			
		mainFrame.setVisible(true);
			}
		});
	}

}
