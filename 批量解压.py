# __coding=utf8__
# /** 作者：zengyanghui **/
import zipfile
import os
import sys

#  reload(sys)
#  sys.setdefaultencoding('gbk')  # 如遇到无法识别中文而报错使用


# 将zip文件解压处理，并放到指定的文件夹里面去

def unzip_file(zip_file_name, destination_path):
    archive = zipfile.ZipFile(zip_file_name, mode='r')
    for file in archive.namelist():
        archive.extract(file, destination_path)
        print(file)
        print(destination_path)
    print(archive)


a = "/Users/maclove/Downloads/2021/typedata/天猫/账单/惠优购天猫旗舰店"  # zipfile 的路径
b = "/Users/maclove/Downloads/2021/typedata/天猫/账单/惠优购天猫旗舰店"  # 解压到路径unzip下


def zipfile_name(file_dir):
    # 读取文件夹下面的文件名.zip
    L = []
    for root, dirs, files in os.walk(file_dir):
        print(root)
        print(dir)
        print(files)
        for file in files:
            if os.path.splitext(file)[1] == '.zip':  # 读取带zip 文件
                L.append(os.path.join(root, file))
                # print(L)
            print(file)
            print(L)
        print(L)
    return L


# 入口函数
def main():
    fn = zipfile_name(a)
    print(fn)
    for file in fn:
        unzip_file(file, b)
        print(file)

def an_garcode(dir_names):
    """anti garbled code"""
    os.chdir(dir_names)


    for temp_name in os.listdir('.'):
        try:
            #使用cp437对文件名进行解码还原
            new_name = temp_name.encode('cp437')
            #win下一般使用的是gbk编码
            new_name = new_name.decode("gbk")
            #对乱码的文件名及文件夹名进行重命名
            os.rename(temp_name, new_name)
            #传回重新编码的文件名给原文件名
            temp_name = new_name
        except:
            #如果已被正确识别为utf8编码时则不需再编码
            pass


        if os.path.isdir(temp_name):
            #对子文件夹进行递归调用
            an_garcode(temp_name)
            #记得返回上级目录
            os.chdir('..')


if __name__ == "__main__":
    # an_garcode(os.getcwd())
    main()
print("done")