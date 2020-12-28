import java.io.InputStream;
import java.util.Scanner;

public class Test {
    public static void main(String[] args) throws ClassNotFoundException {
        boolean flag = true; //标志变量
        while (flag) {
            //预先获取读文件
            //File file = new File("F:\\idea\\自己的代码\\CmdShop\\src\\users.xlsx"); //创建文件对象
            InputStream in1 = Class.forName("Test").getResourceAsStream("/users.xlsx"); //直接使用输入流的方法可移植性更好  而不用FileInputStream的方法
            ReadUserExcel readUser = new ReadUserExcel(); // 创建ReadExcel类的对象
            User users[] = readUser.readExcel(in1);  //read对象调用readExcel方法，读取文件，会返回一个用户类型的数组，并赋值给users数组

            InputStream in2 = Class.forName("Test").getResourceAsStream("/product.xlsx"); //直接使用输入流的方法可移植性更好  而不用FileInputStream的方法
            ReadProductExcel readProduct = new ReadProductExcel(); // 创建ReadExcel类的对象
            Product products[] = readProduct.readExcel(in2);


            //获取用户名和密码
            Scanner sc = new Scanner(System.in);
            System.out.println("请输入用户名：");
            String username = sc.next();
            System.out.println("请输入密码：");
            String password = sc.next();

            //将用户输入的名字和密码与Excel中的作比对
            for (User user:users) {
                if (username.equals(user.getUsername()) && password.equals(user.getPassword())) {
                    System.out.println("登录成功");

                    for(Product product:products){
                        System.out.println("商品ID：" + product.getPID() +
                                           "   商品名称：" + product.getPName() +
                                           "   商品价格: " + product.getPrice() +
                                           "   商品描述: " +product.getDescription()
                        );
                    }
                    flag= false; //当登录成功时变量为假，不再执行while循环
                    break;
                } else {
                    System.out.println("登录失败");
                    break;
                }
            }
        }
    }
}
