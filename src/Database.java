import java.awt.Image;
import java.awt.Toolkit;
import javax.swing.table.DefaultTableModel;
import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.ImageIcon;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */

/**
 *
 * @author Maryann
 */
public class Database extends javax.swing.JFrame {

    /**
     * Creates new form Database
     */
    
    DefaultTableModel model1,model2,model3;
    
    public final void ExportDatabase(){
        try{
            Workbook wb=new XSSFWorkbook();//Creates new workbook
            
            //Customer Details
            
            Sheet sheet1=wb.createSheet("Details");//Creates worksheet 'Details' in the workbook
            Row rowCol1=sheet1.createRow(0);

            //Loop to export header row
            for(int i=0;i<details_table.getColumnCount();i++){
                Cell cell=rowCol1.createCell(i);
                cell.setCellValue(details_table.getColumnName(i));
            }

            for(int j=0;j<details_table.getRowCount();j++){
                Row row1=sheet1.createRow(j+1);//Creates row

                //Loop to assign each table cell to a worksheet cell
                for(int k=0;k<details_table.getColumnCount();k++){
                    Cell cell2=row1.createCell(k);
                    cell2.setCellValue(details_table.getValueAt(j, k).toString());
                    
                }
            }
            
            //Total Transactions
            Sheet sheet2=wb.createSheet("Transactions (TOTAL)");//Creates total transactions worksheet in the workbook
            Row rowCol2=sheet2.createRow(0);

            //Loop to export header row
            for(int l=0;l<totaltrans_table.getColumnCount();l++){
                Cell cell3=rowCol2.createCell(l);
                cell3.setCellValue(totaltrans_table.getColumnName(l));
            }

            for(int m=0;m<totaltrans_table.getRowCount();m++){
                Row row2=sheet2.createRow(m+1);//Creates row

                //Loop to assign each table cell to a worksheet cell
                for(int n=0;n<totaltrans_table.getColumnCount();n++){
                    Cell cell4=row2.createCell(n);
                    cell4.setCellValue(totaltrans_table.getValueAt(m, n).toString());
                    
                }
            }
            
            //Daily Transactions
            Sheet sheet3=wb.createSheet("Transactions (DAY)");//Creates daily transactions worksheet in the workbook
            Row rowCol3=sheet3.createRow(0);

            //Loop to export header row
            for(int o=0;o<dailytrans_table.getColumnCount();o++){
                Cell cell5=rowCol3.createCell(o);
                cell5.setCellValue(dailytrans_table.getColumnName(o));
            }

            for(int p=0;p<dailytrans_table.getRowCount();p++){
                Row row3=sheet3.createRow(p+1);//Creates row

                //Loop to assign each table cell to a worksheet cell
                for(int q=0;q<dailytrans_table.getColumnCount();q++){
                    Cell cell6=row3.createCell(q);
                    cell6.setCellValue(dailytrans_table.getValueAt(p, q).toString());
                    
                }
            }



            wb.write(new FileOutputStream("./Resources\\Database.xlsx"));//Write to xlsx file
            wb.close();



        }catch(IOException e){
            Logger.getLogger(Signup.class.getName()).log(Level.SEVERE, null, e);//Logs error
        }
        
        
    }
    
    public final void ImportDatabase(){
        try{
            Workbook wb=new XSSFWorkbook("./Resources\\Database.xlsx");//Reads xlsx file and genertes a workbook
            
            //Customer Details
            Sheet sheet1=wb.getSheetAt(0);//Gets first worksheet
            Row rowCol=sheet1.getRow(0);//Gets first row (Headers) and assigns it to rowCol
            
            while(details_table.getRowCount()>0){
                model1.removeRow(0);
            }

            //Get number of columns in worksheet to set size of array
            int colnum=rowCol.getLastCellNum()+1;
            String[] values= new String[colnum];

            //Loops through the table starting from the row under headers
            for(int i=1;i<=sheet1.getLastRowNum();i++){
                Row row=sheet1.getRow(i);

                //Loops through each cell in the current row
                for(int j=0;j<row.getLastCellNum();j++){
                    Cell cell=row.getCell(j);

                    //Adds each cell to "values" array
                    try{  
                        values[j]=cell.getStringCellValue();
                    }
                    catch(Exception e){
                        int cellnum=(int)cell.getNumericCellValue();

                        values[j]=String.valueOf(cellnum);
                    }

                }
                model1.addRow(values);//Adds row using "values" array
            }

            details_table.setModel(model1);//Update table model

            
            //Total Transactions
            Sheet sheet2=wb.getSheetAt(1);//Gets second worksheet
            rowCol=sheet2.getRow(0);//Gets first row (Headers) and assigns it to rowCol
            
            while(totaltrans_table.getRowCount()>0){
                model2.removeRow(0);
            }

            //Get number of columns in worksheet to set size of array
            colnum=rowCol.getLastCellNum()+1;
            values= new String[colnum];

            //Loops through the table starting from the row under headers
            for(int k=1;k<=sheet2.getLastRowNum();k++){
                Row row2=sheet2.getRow(k);

                //Loops through each cell in the current row
                for(int l=0;l<row2.getLastCellNum();l++){
                    Cell cell2=row2.getCell(l);

                    //Adds each cell to "values" array
                    try{  
                        values[l]=cell2.getStringCellValue();
                    }
                    catch(Exception e){
                        int cellnum2=(int)cell2.getNumericCellValue();

                        values[l]=String.valueOf(cellnum2);
                    }

                }
                model2.addRow(values);//Adds row using "values" array
            }

            totaltrans_table.setModel(model2);//Update table model

            
            //Daily Transactions
            Sheet sheet3=wb.getSheetAt(2);//Gets third worksheet
            rowCol=sheet3.getRow(0);//Gets first row (Headers) and assigns it to rowCol
            
            while(dailytrans_table.getRowCount()>0){
                model3.removeRow(0);
            }

            //Get number of columns in worksheet to set size of array
            colnum=rowCol.getLastCellNum()+1;
            values= new String[colnum];

            //Loops through the table starting from the row under headers
            for(int m=1;m<=sheet3.getLastRowNum();m++){
                Row row3=sheet3.getRow(m);

                //Loops through each cell in the current row
                for(int n=0;n<row3.getLastCellNum();n++){
                    Cell cell3=row3.getCell(n);

                    //Adds each cell to "values" array
                    try{  
                        values[n]=cell3.getStringCellValue();
                    }
                    catch(Exception e){
                        int cellnum3=(int)cell3.getNumericCellValue();

                        values[n]=String.valueOf(cellnum3);
                    }

                }
                model3.addRow(values);//Adds row using "values" array
            }

            dailytrans_table.setModel(model3);//Update table model

            
            
            wb.close();




        //Exception Handling
        }catch(FileNotFoundException e){
           System.out.println(e);
        }catch(IOException e){
           System.out.println(e);
           
        }
    }
    
    public ImageIcon CreateIcon(JLabel imglabel){
        ImageIcon my_img=new ImageIcon(Toolkit.getDefaultToolkit().getImage("./Resources\\Nexus_Logo.png"));//Take given image and put it in my_img object
        Image img=my_img.getImage();
        img=img.getScaledInstance(imglabel.getWidth(),imglabel.getHeight(),Image.SCALE_SMOOTH);
        ImageIcon i=new ImageIcon(img);
        return i;
    }
    
    
    public Database() {
        initComponents();
        setTitle("Database");
        setDefaultCloseOperation(this.DO_NOTHING_ON_CLOSE);
        
        model1=(DefaultTableModel) details_table.getModel();
        model2=(DefaultTableModel) totaltrans_table.getModel();
        model3=(DefaultTableModel) dailytrans_table.getModel();
        
        ImportDatabase();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        Tables_tabbedpane = new javax.swing.JTabbedPane();
        tab1 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        details_table = new javax.swing.JTable();
        tab2 = new javax.swing.JPanel();
        jScrollPane3 = new javax.swing.JScrollPane();
        totaltrans_table = new javax.swing.JTable();
        tab3 = new javax.swing.JPanel();
        jScrollPane2 = new javax.swing.JScrollPane();
        dailytrans_table = new javax.swing.JTable();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setResizable(false);

        Tables_tabbedpane.setPreferredSize(new java.awt.Dimension(803, 540));
        Tables_tabbedpane.addChangeListener(new javax.swing.event.ChangeListener() {
            public void stateChanged(javax.swing.event.ChangeEvent evt) {
                Tables_tabbedpaneStateChanged(evt);
            }
        });

        details_table.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Name", "AccountNumber", "PinNumber", "ContactNumber", "Balance"
            }
        ) {
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false
            };

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        jScrollPane1.setViewportView(details_table);

        javax.swing.GroupLayout tab1Layout = new javax.swing.GroupLayout(tab1);
        tab1.setLayout(tab1Layout);
        tab1Layout.setHorizontalGroup(
            tab1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(tab1Layout.createSequentialGroup()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 670, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
        );
        tab1Layout.setVerticalGroup(
            tab1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 523, Short.MAX_VALUE)
        );

        Tables_tabbedpane.addTab("Details", tab1);

        totaltrans_table.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "AccountNumber", "TotalWithdraw", "TotalDeposit", "TotalTransfer (Sent)", "TotalTransfer (Received)"
            }
        ) {
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false
            };

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        jScrollPane3.setViewportView(totaltrans_table);
        if (totaltrans_table.getColumnModel().getColumnCount() > 0) {
            totaltrans_table.getColumnModel().getColumn(0).setResizable(false);
            totaltrans_table.getColumnModel().getColumn(1).setResizable(false);
            totaltrans_table.getColumnModel().getColumn(2).setResizable(false);
            totaltrans_table.getColumnModel().getColumn(4).setResizable(false);
        }

        javax.swing.GroupLayout tab2Layout = new javax.swing.GroupLayout(tab2);
        tab2.setLayout(tab2Layout);
        tab2Layout.setHorizontalGroup(
            tab2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(tab2Layout.createSequentialGroup()
                .addComponent(jScrollPane3, javax.swing.GroupLayout.PREFERRED_SIZE, 670, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
        );
        tab2Layout.setVerticalGroup(
            tab2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane3, javax.swing.GroupLayout.DEFAULT_SIZE, 523, Short.MAX_VALUE)
        );

        Tables_tabbedpane.addTab("Transactions (TOTAL)", tab2);

        dailytrans_table.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "AccountNum", "Withdraw", "Deposit", "Transfer (Sent)", "Transfer (Received)"
            }
        ) {
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false
            };

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        jScrollPane2.setViewportView(dailytrans_table);
        if (dailytrans_table.getColumnModel().getColumnCount() > 0) {
            dailytrans_table.getColumnModel().getColumn(0).setResizable(false);
            dailytrans_table.getColumnModel().getColumn(1).setResizable(false);
            dailytrans_table.getColumnModel().getColumn(2).setResizable(false);
            dailytrans_table.getColumnModel().getColumn(3).setResizable(false);
            dailytrans_table.getColumnModel().getColumn(4).setResizable(false);
        }

        javax.swing.GroupLayout tab3Layout = new javax.swing.GroupLayout(tab3);
        tab3.setLayout(tab3Layout);
        tab3Layout.setHorizontalGroup(
            tab3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(tab3Layout.createSequentialGroup()
                .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 670, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
        );
        tab3Layout.setVerticalGroup(
            tab3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 523, Short.MAX_VALUE)
        );

        Tables_tabbedpane.addTab("Transactions (DAY)", tab3);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(Tables_tabbedpane, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 674, javax.swing.GroupLayout.PREFERRED_SIZE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(Tables_tabbedpane, javax.swing.GroupLayout.DEFAULT_SIZE, 550, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void Tables_tabbedpaneStateChanged(javax.swing.event.ChangeEvent evt) {//GEN-FIRST:event_Tables_tabbedpaneStateChanged
        // TODO add your handling code here:
        
        //Confirm manager accnum to access total transaction history
        if(Tables_tabbedpane.getSelectedIndex()==1){
            tab2.setVisible(false);
            jScrollPane1.setVisible(false);
            totaltrans_table.setVisible(false);
            try{
                if(Integer.parseInt(JOptionPane.showInputDialog(this,"Confirm manager account number: "))!=999999){
                    JOptionPane.showMessageDialog(this, "Invalid account number entered");
                    Tables_tabbedpane.setSelectedIndex(0);
                }
                else{
                    tab2.setVisible(true);
                    jScrollPane1.setVisible(true);
                    totaltrans_table.setVisible(true);
                }
            }
            catch(Exception e){
                JOptionPane.showMessageDialog(this, "Invalid account number entered");
                Tables_tabbedpane.setSelectedIndex(0);
            }
        }
    }//GEN-LAST:event_Tables_tabbedpaneStateChanged

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Database.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Database.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Database.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Database.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Database().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JTabbedPane Tables_tabbedpane;
    public javax.swing.JTable dailytrans_table;
    public javax.swing.JTable details_table;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JPanel tab1;
    private javax.swing.JPanel tab2;
    private javax.swing.JPanel tab3;
    public javax.swing.JTable totaltrans_table;
    // End of variables declaration//GEN-END:variables

    
}
