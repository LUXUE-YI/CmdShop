import java.io.File;
import java.util.Scanner;

public class Test {
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);

        System.out.println("请输入用户名：");
        String name = sc.next();

        System.out.println("请输入密码：");
        String passwrod = sc.next();

        //读文件
        File file = new Flie("C:\Users\lenovo\IdeaProjects\CmdShop\src\users.xlsx");
        ReadExcel read = new ReadExcel();
        read.readExcel(file);


    }
}
