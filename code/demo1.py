import os
import zipfile


def walk_filter(path):
    demo = []
    for curDir, dirs, files in os.walk(path):
        for file in files:
            demo.append(os.path.join(curDir, file))
    return demo


def save():
    newf = zipfile.ZipFile("demo/test.docx", 'w', zipfile.ZIP_DEFLATED)  # 创建一个新的docx文件，作为修改后的docx
    for path, dirnames, filenames in os.walk("demo/实验1"):  # 将workfolder文件夹所有的文件压缩至new.docx
        # 去掉目标跟路径，只对目标文件夹下边的文件及文件夹进行压缩
        fpath = path.replace("demo/实验1", '')
        for filename in filenames:
            # print(os.path.join(path, filename))
            newf.write(os.path.join(path, filename), os.path.join(fpath, filename))
    newf.close()


if __name__ == "__main__":
    save()
    path = r"C:\Users\qingfusheng\Desktop\实验2"
    file_paths = walk_filter(path)
    for each in file_paths:
        with open(each, "r", encoding="utf-8") as f:
            content = f.read()
        if "comments" in content:
            print(each)
