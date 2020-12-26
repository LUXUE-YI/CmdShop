import java.io.File;
import java.util.Scanner;

public class Test {
    public static void main(String[] args) {
        //预先获取读文件
        File file = new File("F:\\idea\\自己的代码\\CmdShop\\src\\users.xlsx");
        ReadExcel read = new ReadExcel();
        User users[] = read.readExcel(file);
        System.out.println(users[0].getUsername());
        System.out.println(users[0].getPassword());

        //获取用户名和密码
        Scanner sc = new Scanner(System.in);
        System.out.println("请输入用户名：");
        String username = sc.next();
        System.out.println("请输入密码：");
        String password = sc.next();

        //将用户输入的名字和密码与Excel中的作比对
        for(User user:users){
            if(username.equals(user.getUsername()) && password.equals(user.getPassword())){
                System.out.println("登录成功");
            }
        }
    }
}
