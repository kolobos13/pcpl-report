package ua.org.pycb6a.pcplreport;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Image;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.charset.Charset;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.border.EtchedBorder;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.csvreader.CsvReader;

public class ReportFrame extends JFrame {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	/**
	 * 
	 */

	private JButton buttonConvert;
	private JPanel panelCenter;
	private JTextField textInput;
	private JTextField textOutput;
	private JButton buttonInput;
	private JButton buttonOutput;
	private String textInputDefault = "               [Input File].csv";
	private String textOutputDefault = "               [Output File].xls";
	private File fileInput = null;
	private File fileOutput = null;

	public ReportFrame() {
		super();
		initGUI();
	}

	private void initGUI() {
		try {
			this.setSize(400, 130);
			Image iconReport = Toolkit.getDefaultToolkit().getImage("pcplReport_res/img/report.png");
			this.setIconImage(iconReport);

			buttonConvert = new JButton();
			getContentPane().add(buttonConvert, BorderLayout.SOUTH);
			buttonConvert.setText("Convert File");
			buttonConvert.setPreferredSize(new Dimension(300, 30));
			buttonConvert.addActionListener(new ActionListener() {

				public void actionPerformed(ActionEvent e) {

					ConvertFile();

				}
			});

			panelCenter = new JPanel();
			getContentPane().add(panelCenter, BorderLayout.CENTER);

			panelCenter.setLayout(new FlowLayout());

			textInput = new JTextField();
			panelCenter.add(textInput);
			textInput.setFont(new java.awt.Font("Arial", 0, 12));
			textInput.setPreferredSize(new Dimension(300, 25));
			textInput.setText(textInputDefault);
			textInput.setBorder(BorderFactory
					.createEtchedBorder(EtchedBorder.LOWERED));
			textInput.setEditable(false);

			buttonInput = new JButton();
			panelCenter.add(buttonInput);
			buttonInput.setText("Input");
			buttonInput.setPreferredSize(new Dimension(70, 25));
			buttonInput.addActionListener(new ActionListener() {

				public void actionPerformed(ActionEvent e) {
					setInputText();
				}
			});

			textOutput = new JTextField();
			panelCenter.add(textOutput);
			textOutput.setFont(new java.awt.Font("Arial", 0, 12));
			textOutput.setPreferredSize(new Dimension(300, 25));
			textOutput.setText(textOutputDefault);
			textOutput.setBorder(BorderFactory
					.createEtchedBorder(EtchedBorder.LOWERED));
			textOutput.setEditable(false);

			buttonOutput = new JButton();
			panelCenter.add(buttonOutput);
			buttonOutput.setText("Output");
			buttonOutput.setPreferredSize(new Dimension(70, 25));
			buttonOutput.addActionListener(new ActionListener() {

				public void actionPerformed(ActionEvent e) {
					setOutputText();
				}
			});

		} catch (Exception e) {
			// add your error handling code here
			e.printStackTrace();
		}
	}

	private void setInputText() {
		JFileChooser jInputChooser = new JFileChooser();
		jInputChooser.addChoosableFileFilter(new FileFilter() {

			public boolean accept(File file) {
				String filename = file.getName();
				return filename.endsWith(".csv") || file.isDirectory();
			}

			public String getDescription() {
				return "*.csv";
			}
		});

		try {
			UIManager
					.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
			SwingUtilities.updateComponentTreeUI(jInputChooser);
		} catch (Exception e) {
			e.printStackTrace();
		}
		int result = jInputChooser.showOpenDialog(ReportFrame.this);
		if (result == JFileChooser.APPROVE_OPTION) {
			try {
				textInput.setText(jInputChooser.getSelectedFile().getPath());

			} catch (Exception e) {
				e.printStackTrace();
			}

		}

	}

	private void setOutputText() {
		JFileChooser jOutputChooser = new JFileChooser();
		jOutputChooser.addChoosableFileFilter(new FileFilter() {

			public boolean accept(File f) {
				String fileName = f.getName();
				return fileName.endsWith(".xls") || f.isDirectory();
			}

			public String getDescription() {
				return "*.xls";
			}
		});

		try {
			UIManager
					.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
			SwingUtilities.updateComponentTreeUI(jOutputChooser);
		} catch (Exception e) {
			e.printStackTrace();
		}
		int result = jOutputChooser.showSaveDialog(ReportFrame.this);
		if (result == JFileChooser.APPROVE_OPTION) {
			try {
				String outputTemp = jOutputChooser.getSelectedFile().getPath();
				if (outputTemp.endsWith(".xls"))

					textOutput.setText(jOutputChooser.getSelectedFile()
							.getPath());

				else
					textOutput.setText(jOutputChooser.getSelectedFile()
							.getPath() + ".xls");

				fileOutput = new File(textOutput.getText());
			} catch (Exception e) {
				e.printStackTrace();
			}

		}

	}

	private void ConvertFile() {
		fileInput = new File(textInput.getText());
		if (fileInput.exists()) {
			if (!textOutput.getText().equals(textOutputDefault)) {
				try {

					FileInputStream fis = new FileInputStream(fileInput);
					CsvReader products = new CsvReader(fis,
							Charset.forName("Cp1251"));
					products.skipLine();
					products.readHeaders();

					POIFSFileSystem fisTemp = new POIFSFileSystem(
							new FileInputStream("pcplReport_res/template/template.xls"));
					HSSFWorkbook hwb = new HSSFWorkbook(fisTemp, true);
					HSSFSheet sheet = hwb.getSheet("Детальный отчет");

					int rowIn = 1;
					while (products.readRecord()) {
						String Time = products.get("Time");
						String User = products.get("User");
						String Pages = products.get("Pages");
						int iPages = Integer.parseInt(Pages);

						String Copies = products.get("Copies");
						int iCopies = Integer.parseInt(Copies);

						int Count = Integer.parseInt(Pages)
								* Integer.parseInt(Copies);

						String Printer = products.get("Printer");
						String Document = products.get("Document Name");
						String Client = products.get("Client");
						// String Paper = products.get("Paper");
						// String Language = products.get("Language");
						// String Height = products.get("Height");
						// String Width = products.get("Width");
						// String Duplex = products.get("Duplex");
						// String Grayscale = products.get("Grayscale");
						String Size = products.get("Size");

						HSSFRow row = sheet.createRow(rowIn++);

						HSSFCell cell = row.createCell(0);
						cell.setCellValue(Time);
						cell = row.createCell(1);
						cell.setCellValue(Printer);
						cell = row.createCell(2);
						cell.setCellValue(Client);
						cell = row.createCell(3);
						cell.setCellValue(User);
						cell = row.createCell(4);
						cell.setCellValue(Document);
						cell = row.createCell(5);
						cell.setCellValue(iPages);
						cell = row.createCell(6);
						cell.setCellValue(iCopies);
						cell = row.createCell(7);
						cell.setCellValue(Count);
						cell = row.createCell(8);
						cell.setCellValue(Size);

					}

					products.close();

					FileOutputStream fos = new FileOutputStream(fileOutput);
					hwb.write(fos);
					fis.close();
					fos.close();

					JOptionPane.showMessageDialog(ReportFrame.this, fileOutput
							+ " создан.");
					Runtime.getRuntime().exec("cmd.exe /r " + fileOutput);

				}

				catch (Exception e) {

					JOptionPane.showMessageDialog(ReportFrame.this,
							"Что-то не так.");

				}
			} else
				JOptionPane.showMessageDialog(ReportFrame.this,
						"Выберите XLS-файл.");

		} else
			JOptionPane.showMessageDialog(ReportFrame.this,
				  "Выберите CSV-файл.");
	}

}
