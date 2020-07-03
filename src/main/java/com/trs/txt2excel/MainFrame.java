package com.trs.txt2excel;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.border.TitledBorder;

import com.trs.txt2excel.parser.TxtLineParser;
import com.trs.txt2excel.util.FileUtil;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class MainFrame extends JFrame {

	private static final long serialVersionUID = 1L;

	private JPanel contentPane;

	private JTextField selectedFileName;
	private JTextArea previewTextArea;
	private JTextArea resultTextArea;

	private JButton openExportFileBtn;

	private JProgressBar progressBar;

	private String firstLine;
	private List<String> txtLines;
	private File openFile;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		SwingUtilities.invokeLater(new Runnable() {
			@Override
			public void run() {
				try {
					MainFrame frame = new MainFrame();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public MainFrame() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		// 显示在屏幕中间
		setFrameToCenter(600, 600);
		setResizable(false);

		contentPane = new JPanel();
		setContentPane(contentPane);
		contentPane.setLayout(null);

		JPanel chooseFilePanel = new JPanel();
		chooseFilePanel.setBorder(new TitledBorder(null, "TXT文本", TitledBorder.LEFT, TitledBorder.TOP, null, null));
		chooseFilePanel.setBounds(10, 10, 574, 311);
		contentPane.add(chooseFilePanel);
		chooseFilePanel.setLayout(null);

		JButton chooseBtn = new JButton("选择文件");
		chooseBtn.setBounds(10, 27, 103, 23);
		chooseFilePanel.add(chooseBtn);

		selectedFileName = new JTextField();
		selectedFileName.setBounds(123, 27, 441, 23);
		chooseFilePanel.add(selectedFileName);
		selectedFileName.setEditable(false);
		selectedFileName.setColumns(10);

		final JTextArea textArea = new JTextArea();
		textArea.setEditable(false);
		JScrollPane scrollPane = new JScrollPane(textArea);
		scrollPane.setBounds(10, 60, 554, 46);
		chooseFilePanel.add(scrollPane);
		final JButton exportBtn = new JButton("导出Excel");
		exportBtn.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				SwingWorker<String, String> worker = new SwingWorker<String, String>() {
					@Override
					protected String doInBackground() throws Exception {
						String exportFile = openFile.getAbsolutePath() + ".xlsx";
						resultTextArea.append("开始导出Excel文件：" + exportFile + "\n");
						try (Workbook wb = new SXSSFWorkbook();
								OutputStream stream = new FileOutputStream(exportFile);) {
							Sheet sheet = wb.createSheet("sheet1");
							int total = txtLines.size();
							progressBar.setMaximum(total);
							for (int i = 0; i < total;) {
								Row row = sheet.createRow(i);
								int cellNum = 0;
								for (String value : TxtLineParser.readLine(txtLines.get(i))) {
									row.createCell(cellNum++).setCellValue(value);
								}
								i++;
								if (total < 100 || i % (total / 100) == 0) {
									progressBar.setValue(i + 1);
								}
							}
							wb.write(stream);
							progressBar.setValue(total);
							openExportFileBtn.setEnabled(true);
						} catch (Exception e2) {
							resultTextArea.append("导出Excel文件失败\n" + e2);
						}
						resultTextArea.append("导出Excel文件完成\n");
						return "OK";
					}
				};
				worker.execute();
			}
		});
		exportBtn.setEnabled(false);
		exportBtn.setBounds(10, 331, 93, 23);
		contentPane.add(exportBtn);
		textArea.setEditable(false);

		final JButton previewBtn = new JButton("解析预览");
		previewBtn.setBounds(10, 116, 93, 23);
		chooseFilePanel.add(previewBtn);
		previewBtn.setEnabled(false);
		previewBtn.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				previewTextArea.setText("");
				for (String line : TxtLineParser.readLine(firstLine)) {
					previewTextArea.append(line);
					previewTextArea.append("\n");
				}
			}
		});

		previewTextArea = new JTextArea();
		JScrollPane scrollPane1 = new JScrollPane(previewTextArea);
		scrollPane1.setBounds(10, 149, 554, 152);
		chooseFilePanel.add(scrollPane1);

		openExportFileBtn = new JButton("打开导出文件所在目录");
		openExportFileBtn.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				SwingUtilities.invokeLater(new Runnable() {
					@Override
					public void run() {
						if (openFile != null) {
							try {
								resultTextArea.append("打开导出文件所在目录:" + openFile.getParentFile() + "\n");
								Desktop.getDesktop().browse(openFile.getParentFile().toURI());
							} catch (IOException e1) {
								resultTextArea.append("打开导出文件所在目录失败\n" + e1);
							}
						}
					}
				});
			}
		});
		openExportFileBtn.setEnabled(false);
		openExportFileBtn.setBounds(113, 331, 180, 23);
		contentPane.add(openExportFileBtn);

		JPanel panel = new JPanel();
		panel.setBorder(new TitledBorder(null, "日志输出", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel.setBounds(10, 364, 574, 172);
		contentPane.add(panel);
		panel.setLayout(new BorderLayout(0, 0));

		resultTextArea = new JTextArea();
		resultTextArea.setSize(279, 186);
		JScrollPane scrollPane2 = new JScrollPane(resultTextArea);
		panel.add(scrollPane2);

		progressBar = new JProgressBar();
		progressBar.setBounds(384, 539, 200, 23);
		progressBar.setStringPainted(true);
		progressBar.setForeground(Color.GREEN);
		contentPane.add(progressBar);

		chooseBtn.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// 打开当前目录
				JFileChooser jfc = new JFileChooser(openFile != null ? openFile : new File("."));
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc.setDialogTitle("选择要转为Excel的文本文件");
				if (JFileChooser.APPROVE_OPTION == jfc.showOpenDialog(null)) {
					openFile = jfc.getSelectedFile();
					selectedFileName.setText(openFile.getAbsolutePath());
					resultTextArea.append("输入文件：" + openFile.getAbsolutePath() + "\n");
					try {
						txtLines = FileUtil.loadText(openFile, "UTF-8");
						resultTextArea.append("一共：" + txtLines.size() + " 行\n");
						if (!txtLines.isEmpty()) {
							firstLine = txtLines.get(0);
						}
						textArea.setText(firstLine);
						previewBtn.setEnabled(true);
						exportBtn.setEnabled(true);
						openExportFileBtn.setEnabled(false);
						progressBar.setValue(0);
					} catch (IOException e1) {
						resultTextArea.append("文件读取错误：\n" + e1 + "\n");
						firstLine = null;
					}
				}

			}
		});
	}

	/**
	 * 设置窗口显示再屏幕中间
	 * 
	 * @param width  窗口宽
	 * @param height 窗口高
	 * @since hanchao @ 2020年7月3日
	 */
	private void setFrameToCenter(int width, int height) {
		Toolkit kit = Toolkit.getDefaultToolkit(); // 定义工具包
		Dimension screenSize = kit.getScreenSize(); // 获取屏幕的尺寸
		int x = (screenSize.width - width) / 2; // 获取屏幕的宽
		int y = (screenSize.height - height) / 2; // 获取屏幕的高
		setBounds(x, y, width, height);
	}
}
