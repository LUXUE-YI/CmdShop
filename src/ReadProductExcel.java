import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;

public class ReadProductExcel {
    /*
    readExcel是什么方法？成员方法
    readExcel的返回值是一个用户类型的一维数组，如String[]，数组中存储一个个字符串，而Users[]存储一个个用户，用户有自己的名称、密码、地址、电话
    readExcel的的形参是File类型的文件，调用方法时实参为一个文件类型的文件，才能解析出对应的文件内容
     */
    public Product[] getAllProduct(InputStream in) {
        Product products[] = null;  //定义一个user类型的一维数组users，其中存储一个个用户，初值为空，无用户
        try {
            XSSFWorkbook xw = new XSSFWorkbook(in);  //创建XSSFWorkbook类的对象xw
            XSSFSheet xs = xw.getSheetAt(0);  //xw对象调用getSheetAt()方法，并赋值给xs，以获取Excel文件中的第一个表格（一个文件中有多个表格）
            products = new Product[xs.getLastRowNum()]; //xs.getLastRowNum()获取表格中的存有信息的总行数，并给users数组分配空间，相当于users = new User[3]

            //从表格的第二行（因为第一行是表头标题，没有真正存储用户信息）开始对存有用户信息的每一行进行遍历，获取每行的信息
            for (int j = 1; j <= xs.getLastRowNum(); j++) {
                XSSFRow row = xs.getRow(j); //调用方法获取表格对象xs的第二行，赋给XSSFRow类型的row
                Product product = new Product();   //创建一个User类的对象，实例化

                //从每一行的第一个单元格开始，获取该行每个单元格的信息
                for (int k = 0; k <= row.getLastCellNum(); k++) {
                    XSSFCell cell = row.getCell(k); //获取该行row中的单元格，赋值给cell
                    if (cell == null)   //如果单元格为空，则跳出本次循环，执行下一次循环
                        continue;
                    if (k == 0) {    //如果单元格不为空，并且是该行的第一个单元格，则调用方法获取这个单元格的值，并给用户的名称属性赋值
                        product.setPID(this.getValue(cell));
                    }
                    else if (k == 1) {   //如果单元格不为空，并且是该行的第二个单元格，则调用方法获取这个单元格的值，并给用户的密码属性赋值
                        product.setPName(this.getValue(cell));
                    }
                    else if (k == 2) {   //如果单元格不为空，并且是该行的第三个单元格，则调用方法获取这个单元格的值，并给用户的地址属性赋值
                        product.setPrice(this.getValue(cell));
                    }
                    else if (k == 3) {   //如果单元格不为空，并且是该行的第一个单元格，则调用方法获取这个单元格的值，并给用户的电话号码属性赋值
                        product.setDescription(this.getValue(cell));
                    }
                    products[j-1]=product;  //将已有信息的用户存储到users数组中，j从1开始，但是存储到数组时应该存到users[0],
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return products;  //将存有用户信息的数组返回
    }

    public Product getProductByID(String id,InputStream in) {
        try {
            XSSFWorkbook xw = new XSSFWorkbook(in);  //创建XSSFWorkbook类的对象xw
            XSSFSheet xs = xw.getSheetAt(0);  //xw对象调用getSheetAt()方法，并赋值给xs，以获取Excel文件中的第一个表格（一个文件中有多个表格）

            //从表格的第二行（因为第一行是表头标题，没有真正存储用户信息）开始对存有用户信息的每一行进行遍历，获取每行的信息
            for (int j = 1; j <= xs.getLastRowNum(); j++) {
                XSSFRow row = xs.getRow(j); //调用方法获取表格对象xs的第二行，赋给XSSFRow类型的row
                Product product = new Product();   //创建一个User类的对象，实例化

                //从每一行的第一个单元格开始，获取该行每个单元格的信息
                for (int k = 0; k <= row.getLastCellNum(); k++) {
                    XSSFCell cell = row.getCell(k); //获取该行row中的单元格，赋值给cell
                    if (cell == null)   //如果单元格为空，则跳出本次循环，执行下一次循环
                        continue;
                    if (k == 0) {    //如果单元格不为空，并且是该行的第一个单元格，则调用方法获取这个单元格的值，并给用户的名称属性赋值
                        product.setPID(this.getValue(cell));
                    }
                    else if (k == 1) {   //如果单元格不为空，并且是该行的第二个单元格，则调用方法获取这个单元格的值，并给用户的密码属性赋值
                        product.setPName(this.getValue(cell));
                    }
                    else if (k == 2) {   //如果单元格不为空，并且是该行的第三个单元格，则调用方法获取这个单元格的值，并给用户的地址属性赋值
                        product.setPrice(this.getValue(cell));
                    }
                    else if (k == 3) {   //如果单元格不为空，并且是该行的第一个单元格，则调用方法获取这个单元格的值，并给用户的电话号码属性赋值
                        product.setDescription(this.getValue(cell));
                    }
                }
                if(id.equals(product.getPID())){
                    return product;
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;  //将存有用户信息的数组返回
    }
    /*
    ReadExcel的另一个成员方法，用于获取单元格中的值
    */
    private String getValue(XSSFCell cell) {
        String value;
        CellType type = cell.getCellTypeEnum();

        DecimalFormat decimalFormat = new DecimalFormat("#");

        switch (type) {
            case STRING:
                value = cell.getStringCellValue();
                break;
            case BLANK:
                value = "";
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue() + "";
                break;
            case NUMERIC:
                //value = cell.getNumericCellValue() + "";
                //double类型的cell（如120.0）和一个空字符串相连，最终为转成字符串（如转成字符串120.0）
                value = decimalFormat.format(cell.getNumericCellValue());
                break;
            case FORMULA:
                value = cell.getCellFormula();
                break;
            case ERROR:
                value = "非法字符";
                break;
            default:
                value = "";
                break;
        }
        return value;
    }
}