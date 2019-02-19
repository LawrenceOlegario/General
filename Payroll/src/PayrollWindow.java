import java.awt.*;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import javax.swing.JSpinner;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Set;
import java.util.TreeMap;
import java.util.logging.Logger;

import javax.swing.JComboBox;

public class PayrollWindow extends JFrame {

	private JPanel contentPane;
	private JTable table;
	private JTextField NameTF;
	private JTextField JobTF;
	private JTextField AttendanceTF;
	private JTextField WageTF;
	private JTextField AdditionTF;
	private JTextField DeductionTF;
	private JTextField AdditionTypeTF;
	private JTextField DeductionTypeTF;
	private JTextField project;

	/**
	 * Launch the application.
	 */
	
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					PayrollWindow frame = new PayrollWindow();
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
	DefaultTableModel model = new DefaultTableModel();
	
	private String getCellValue(int x, int y) {
		
		return model.getValueAt(x,y).toString();
	}
	
	public PayrollWindow() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 952, 453);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(10, 280, 916, 124);
		contentPane.add(scrollPane);
		
		table = new JTable();
		scrollPane.setViewportView(table);
		Object[] columns = {"Name", "Job", "Pay", "Attendace", "Additions", "Additions Type", "Deductions", "Deductions Type", "Pay"};
		model.setColumnIdentifiers(columns);
		table.setModel(model);
		
		
		
		JLabel lblNewLabel = new JLabel("Name");
		lblNewLabel.setBounds(10, 70, 46, 14);
		contentPane.add(lblNewLabel);
		
		JLabel lblNewLabel_1 = new JLabel("Job");
		lblNewLabel_1.setBounds(10, 95, 46, 14);
		contentPane.add(lblNewLabel_1);
		
		JLabel lblNewLabel_2 = new JLabel("Wage");
		lblNewLabel_2.setBounds(10, 120, 46, 14);
		contentPane.add(lblNewLabel_2);
		
		JLabel lblNewLabel_3 = new JLabel("Attendance");
		lblNewLabel_3.setBounds(10, 145, 66, 14);
		contentPane.add(lblNewLabel_3);
		
		JLabel lblNewLabel_4 = new JLabel("Addition");
		lblNewLabel_4.setBounds(10, 170, 46, 14);
		contentPane.add(lblNewLabel_4);
		
		JLabel lblNewLabel_5 = new JLabel("Deduction");
		lblNewLabel_5.setBounds(10, 195, 66, 14);
		contentPane.add(lblNewLabel_5);
		
		JLabel lblNewLabel_6 = new JLabel("Type");
		lblNewLabel_6.setBounds(412, 170, 46, 14);
		contentPane.add(lblNewLabel_6);
		
		JLabel lblNewLabel_7 = new JLabel("Type");
		lblNewLabel_7.setBounds(412, 195, 46, 14);
		contentPane.add(lblNewLabel_7);
		
		NameTF = new JTextField();
		NameTF.setBounds(86, 67, 239, 20);
		contentPane.add(NameTF);
		NameTF.setColumns(10);
		
		JobTF = new JTextField();
		JobTF.setBounds(86, 92, 239, 20);
		contentPane.add(JobTF);
		JobTF.setColumns(10);
		
		AttendanceTF = new JTextField();
		AttendanceTF.addKeyListener(new KeyAdapter() {
			@Override
			public void keyTyped(java.awt.event.KeyEvent evt) {
				char c=evt.getKeyChar();
				if (Character.isLetter(c) &&! evt.isAltDown()){
					evt.consume();
				}
			}
		});
		
		AttendanceTF.setBounds(86, 142, 86, 20);
		contentPane.add(AttendanceTF);
		AttendanceTF.setColumns(10);
		
		WageTF = new JTextField();
		WageTF.addKeyListener(new KeyAdapter() {
			@Override
			public void keyTyped(java.awt.event.KeyEvent evt) {
				char c=evt.getKeyChar();
				if (Character.isLetter(c) &&! evt.isAltDown()){
					evt.consume();
				}
			}
		});
		
		WageTF.setBounds(86, 117, 86, 20);
		contentPane.add(WageTF);
		WageTF.setColumns(10);
		
		AdditionTF = new JTextField();
		AdditionTF.addKeyListener(new KeyAdapter() {
			@Override
			public void keyTyped(java.awt.event.KeyEvent evt) {
				char c=evt.getKeyChar();
				if (Character.isLetter(c) &&! evt.isAltDown()){
					evt.consume();
				}
			}
		});
		
		AdditionTF.setBounds(86, 167, 86, 20);
		contentPane.add(AdditionTF);
		AdditionTF.setColumns(10);
		
		DeductionTF = new JTextField();
		DeductionTF.addKeyListener(new KeyAdapter() {
			@Override
			public void keyTyped(java.awt.event.KeyEvent evt) {
				char c=evt.getKeyChar();
				if (Character.isLetter(c) &&! evt.isAltDown()){
					evt.consume();
				}
			}
		});
		
		DeductionTF.setBounds(86, 192, 86, 20);
		contentPane.add(DeductionTF);
		DeductionTF.setColumns(10);
		
		AdditionTypeTF = new JTextField();
		AdditionTypeTF.setBounds(468, 167, 151, 20);
		contentPane.add(AdditionTypeTF);
		AdditionTypeTF.setColumns(10);
		
		DeductionTypeTF = new JTextField();
		DeductionTypeTF.setBounds(468, 192, 151, 20);
		contentPane.add(DeductionTypeTF);
		DeductionTypeTF.setColumns(10);
		
		JButton btnNewButton = new JButton("Clear Entries");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				NameTF.setText(null);
				JobTF.setText(null);
				WageTF.setText(null);
				AttendanceTF.setText(null);
				AdditionTF.setText(null);
				DeductionTF.setText(null);
				AdditionTypeTF.setText(null);
				DeductionTypeTF.setText(null);
			}
		});
		btnNewButton.setBounds(812, 111, 114, 23);
		contentPane.add(btnNewButton);
		
		JButton btnNewButton_1 = new JButton("Clear Table");
		btnNewButton_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				model.setRowCount(0);
			}
		});
		btnNewButton_1.setBounds(812, 141, 114, 23);
		contentPane.add(btnNewButton_1);
		
		JButton btnNewButton_2 = new JButton("Exit");
		btnNewButton_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				System.exit(0);
			}
		});
		btnNewButton_2.setBounds(812, 246, 114, 23);
		contentPane.add(btnNewButton_2);
		
		JButton btnNewButton_3 = new JButton("Add to Table");
		btnNewButton_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				
				Double Wage = Double.parseDouble(WageTF.getText());
				Double Attendance = Double.parseDouble(AttendanceTF.getText());
				Double Addition = Double.parseDouble(AdditionTF.getText());
				Double Deduction = Double.parseDouble(DeductionTF.getText());
				
				Double Pay = (Wage*Attendance)+Addition-Deduction;
				
				model.addRow(new Object[] {NameTF.getText(), JobTF.getText(), WageTF.getText(), AttendanceTF.getText(),
							AdditionTF.getText(), AdditionTypeTF.getText(), DeductionTF.getText(), DeductionTypeTF.getText(),Pay});
				NameTF.setText(null);
				JobTF.setText(null);
				WageTF.setText(null);
				AttendanceTF.setText(null);
				AdditionTF.setText(null);
				DeductionTF.setText(null);
				AdditionTypeTF.setText(null);
				DeductionTypeTF.setText(null);
			}
		});
		btnNewButton_3.setBounds(812, 80, 114, 23);
		contentPane.add(btnNewButton_3);
		
		JLabel lblNewLabel_8 = new JLabel("Payroll System");
		lblNewLabel_8.setFont(new Font("Tahoma", Font.PLAIN, 24));
		lblNewLabel_8.setBounds(10, 11, 178, 37);
		contentPane.add(lblNewLabel_8);
					
		

		String [] day = {"Day","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31"};
		String [] month = {"Month","1","2","3","4","5","6","7","8","9","10","11","12"};
		String [] year = {"Year","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30"};
		
		JComboBox fromDay = new JComboBox(day);
		fromDay.setBounds(491, 50, 70, 20);
		contentPane.add(fromDay);
		
		JComboBox fromMonth = new JComboBox(month);
		fromMonth.setBounds(571, 50, 70, 20);
		contentPane.add(fromMonth);
		
		JComboBox fromYear = new JComboBox(year);
		fromYear.setBounds(651, 50, 70, 20);
		contentPane.add(fromYear);
		
		JComboBox toDay = new JComboBox(day);
		toDay.setBounds(491, 92, 70, 20);
		contentPane.add(toDay);
		
		JComboBox toMonth = new JComboBox(month);
		toMonth.setBounds(571, 92, 70, 20);
		contentPane.add(toMonth);
		
		JComboBox toYear = new JComboBox(year);
		toYear.setBounds(651, 92, 70, 20);
		contentPane.add(toYear);
	
		JButton btnNewButton_4 = new JButton("Save Table");
		btnNewButton_4.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				String filename = project.getText()+"   .   "+(String)fromDay.getSelectedItem()+"/"+(String)fromMonth.getSelectedItem()+"/"+(String)fromYear.getSelectedItem()
				+" - "+"/"+(String)toDay.getSelectedItem()+"/"+(String)toMonth.getSelectedItem()+"/"+(String)toYear.getSelectedItem();
				
				XSSFWorkbook wb = new XSSFWorkbook();
				XSSFSheet ws = wb.createSheet();
				
				TreeMap<String, Object[]> data = new TreeMap<>();
				
				data.put("-1", new Object[] {model.getColumnName(0),model.getColumnName(1),model.getColumnName(2),model.getColumnName(3)
						,model.getColumnName(4),model.getColumnName(5),model.getColumnName(6),model.getColumnName(7),model.getColumnName(8)});
				
				for (int i=0;i<model.getRowCount();i++) {
					data.put(Integer.toString(i),new Object[] {getCellValue(i,0),getCellValue(i,1),getCellValue(i,2),getCellValue(i,3)
							,getCellValue(i,4),getCellValue(i,5),getCellValue(i,6),getCellValue(i,7),getCellValue(i,8)});
				}
				Set<String> ids=data.keySet();
				XSSFRow row;
				int rowID=0;
				
				for (String key:ids) {
					row=ws.createRow(rowID++);
					Object[] value=data.get(key);
					
					int cellID=0;
					for (Object o: value) {
						Cell cell = row.createCell(cellID++);
						cell.setCellValue(o.toString());
					}
				}
				try {
					FileOutputStream fos = new FileOutputStream(new File ("C:\\Users\\Lawrence\\Desktop\\"+filename+".xlsx"));
					wb.write(fos);
					fos.close();
					
				}catch(Exception e1) {
					System.out.println("Uh-Oh...");
				}
				
			}
		});
		
		
		btnNewButton_4.setBounds(812, 170, 114, 23);
		contentPane.add(btnNewButton_4);
		
		JLabel label = new JLabel("Project Name");
		label.setBounds(412, 28, 90, 14);
		contentPane.add(label);
		
		project = new JTextField();
		project.setColumns(10);
		project.setBounds(512, 25, 158, 20);
		contentPane.add(project);
		
		JLabel label_1 = new JLabel("Date");
		label_1.setBounds(412, 53, 46, 14);
		contentPane.add(label_1);
		
		JLabel label_2 = new JLabel("From");
		label_2.setBounds(456, 53, 46, 14);
		contentPane.add(label_2);
		
		JLabel label_3 = new JLabel("To");
		label_3.setBounds(456, 92, 46, 14);
		contentPane.add(label_3);
		
	}
}
