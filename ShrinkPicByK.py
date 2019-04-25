# utf 8
# python 3
# Windows 10

########################################## 引入模块的段落

import time                     # 用于计时，延迟等。

import sys                      # 用于计算行号等。

import importlib                # 用于动态加载模块。

import os                       # 用于执行系统命令，并加载模块的命令，以及切换文件夹等。

def import_mod(module_local_name, module_remote_name = ""):
    if module_remote_name == "":
        module_remote_name = module_local_name
        
    try:
        ## module_object = import_module(module_name, class_name)
        importlib.import_module(module_local_name)
    except ImportError as err:

        print(time.strftime('%Y%m%d_%H%M%S',time.localtime()) + " 行号 " + str(sys._getframe().f_lineno) + \
        "错误：" + str(err))    

        print(time.strftime('%Y%m%d_%H%M%S',time.localtime()) + " 行号 " + str(sys._getframe().f_lineno) + \
       " 试图加载模块时失败，再试远程库名 = " + module_remote_name)

        os.system("python -m pip install " + module_remote_name)

        importlib.import_module(module_local_name)

# end of import module ########################################################33
  
import winshell                        # 建立快捷方式用。

from win32com.client import Dispatch   # 创建快捷方式要用到。

import string                          # 做随机数字串时要用到。

from random import choice              # 做随机运算的函数。

# import_mod("numpy")  

from pathlib import Path               # 测试文件及路径是否存在.

import logging                         # 日志文件用.

import_mod("PIL", "Pillow")

from PIL import Image                  # 处理图片时要用到.

from shutil import copyfile            # 文件拷贝时要用到.

import shutil                          # 拷贝文件用到.

# import_mod("winshell")               # 加载模块，用于快捷方式的建立。
        
# 1. 读取命令行参数
# 1b. 如果没有命令行参数，则转入建立“发送到”的快捷方式。
# 2. 建立日志系统
# 3. 读取源图片文件的参数：宽、高、质量。
# 4. 设定缩放比例的起始值、压缩质量的起始值。
# 5. 进行resize，并再做压缩quality
# 6. 验证目标图片文件的大小，符合要求，予以存放。
# 7. 结束。

strAppVersion = "V 1.0.4"
strAppHomepage = "https://github.com/evering7/ShrinkPict"
print("压缩图片，以符合网站上传的要求 by Python 版本号：" + strAppVersion)
print("项目主页：" + strAppHomepage)
print("作者：福建莆田 李剑飞 13799001059@139.com")
print("代码的最后修改日期：2019.4.25")
time.sleep(1)  # 停留一会儿。

def print_time_lineno(strPrint, linenumber = 0):
    if linenumber == 0:
        strLineNumber = str(sys._getframe().f_lineno)
    else:
        strLineNumber = str(linenumber)

    print(time.strftime('%Y%m%d %H%M%S', time.localtime()) + " 行号 " + strLineNumber + " " + strPrint)
    # end of print_time_lineno
##########################################################################################################

# 读取命令行上的所有源文件路径名参数
sourceFiles = sys.argv[1:]

# 1. 读取命令行参数
# 1b. 如果没有命令行参数，则转入建立“发送到”的快捷方式。
if len(sourceFiles) == 0:
    # 此处，开始建立快捷方式。
    print_time_lineno("进入建立快捷方式的代码分支。", sys._getframe().f_lineno)
    
    strSendToFolder = winshell.sendto()
    print_time_lineno("取得发送到文件夹的全路径名 = " + strSendToFolder, sys._getframe().f_lineno)
    
    # 准备创建快捷方式
    shell = Dispatch('WScript.Shell')
    
    strShrinkPict_CmdFullPath = os.path.splitext(os.path.realpath(__file__))[0] + ".cmd"
    print_time_lineno("Command批处理脚本的全路径 = " + strShrinkPict_CmdFullPath, sys._getframe().f_lineno)
    
    strRandom = ''.join(choice(string.ascii_uppercase + string.digits) for i in range(8))
    print_time_lineno("随机字串 = " + strRandom, sys._getframe().f_lineno)
    
    strShortCut_Location = strSendToFolder + "\\压缩图片到上限尺寸之内 " + strRandom + ".lnk"
    print_time_lineno("快捷方式的存放路径全名 = " + strShortCut_Location, sys._getframe().f_lineno)
    
    shortcut = shell.CreateShortCut(strShortCut_Location)
    shortcut.Targetpath = strShrinkPict_CmdFullPath
    
    strScriptContainingFolder = os.path.dirname(__file__)
    print_time_lineno("快捷方式的工作文件夹 = " + strScriptContainingFolder)
    
    shortcut.WorkingDirectory = strScriptContainingFolder
    
    icon_path = strScriptContainingFolder + "\\icons8-edit-image-48.ico"
    print_time_lineno("图标文件的路径 = " + icon_path)
    
    shortcut.IconLocation = icon_path
    shortcut.save()
    
    #完成安装快捷方式的操作，退出
    print_time_lineno("完成安装快捷方式的操作，退出", sys._getframe().f_lineno)
    quit()
    
existSourceFiles = []

for iFile in range(len(sourceFiles)):
    # 开始循环检查.
    currentFile = Path(sourceFiles[iFile])
    
    if currentFile.is_file() and os.path.exists(sourceFiles[iFile]):
        print_time_lineno("文件存在 " + sourceFiles[iFile], sys._getframe().f_lineno)
        existSourceFiles += [sourceFiles[iFile]]
    else:
        print_time_lineno("文件不存在! = " + sourceFiles[iFile] )

print_time_lineno("真正存在的文件列表 = " + str(existSourceFiles))

if len(existSourceFiles) == 0:
    print_time_lineno("所输入的文件参数均不存在,退出.")
    quit()

    
def makeFolderIfNotExist_WithoutFormalLog(folderPath):
    if not os.path.exists(folderPath):
        os.makedirs(folderPath)
        print_time_lineno("创建了文件夹 %s " % folderPath)
        time.sleep(1)
    if not os.path.exists(folderPath):
        print_time_lineno("实在创建不了该文件夹,程序退出.")
        quit()
    
# end of makeFolderIfNotExist_WithoutFormalLog ##################################3
    
# 设定相关参数.

# 2.1 日志文件夹
LogFileNamePrefix = "ShrinkPict"
LogFolder = os.path.join(os.path.dirname(os.path.realpath(__file__)), "LogFolder")
makeFolderIfNotExist_WithoutFormalLog(LogFolder)

# 这个文件夹要创立一下.
print_time_lineno("当前的日志文件夹是 = " + LogFolder)

# 2.2 临时文件夹

TempFolder = os.path.join(os.path.dirname(os.path.realpath(__file__)), "TempFolder")
makeFolderIfNotExist_WithoutFormalLog(TempFolder)
print_time_lineno("当前的临时文件夹是 = " + TempFolder)

# 2.3 文件大小参数.
FileSizeK_UpLimit = int(input("请输入文件尺寸上限(以K为单位) = "))

# 防止误输入
if FileSizeK_UpLimit < 10:
    FileSizeK_UpLimit = 10

realFileSizeInBytes_UpLimit = FileSizeK_UpLimit * 1000



  
# 2.4 建立日志系统
# 2019.4.25
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# 创建了句柄,以便把日志数据导到文件中.
strFixedTime = time.strftime('%Y%m%d %H%M%S', time.localtime())

# 构造日志文件名
strLogFileName = LogFolder + "\\" + LogFileNamePrefix + " " + str(FileSizeK_UpLimit) + "K " + strFixedTime + '.log'

print_time_lineno("当前日志文件的全路径名 = " + strLogFileName , sys._getframe().f_lineno)

logger_handler = logging.FileHandler(strLogFileName)
logger_handler.setLevel(logging.DEBUG)

# 创设一个格式
logger_formatter = logging.Formatter('%(asctime)s %(name)s - %(levelname)s - %(message)s')

# 把格式加入日志文件的句柄
logger_handler.setFormatter(logger_formatter)

# 把句柄加入Logger
logger.addHandler(logger_handler)

# 欢呼:完成日志文件的配置
logger.info("完成日志文件的配置")

# 2b. 读取文件大小的目标上限. ok
# 此处,定义好print_log()

def print_log(strPrint, lineno = 0):                # 开始定义最常用的记录工具例程
    if lineno <= 0:
        lineno = sys._getframe().f_lineno
        
    logger.info("行号 " + str(lineno) + " " + strPrint)
    print(time.strftime('%Y%m%d %H%M%S', time.localtime()) + " 行号 " + str(lineno) + " " + strPrint)
    
    

# end of print_log  ##############################################################

print_log("下面将读取源图片文件的参数:宽,高,质量")

def FileExist(filePath):                     ######################################
    print_log("检验文件是否存在 = " + filePath, sys._getframe().f_lineno)
    myFile = Path(filePath)
    if myFile.is_file():
        print_log("文件存在,放心", sys._getframe().f_lineno)
        return True
    else:
        print_log("意外:文件不存在", sys._getframe().f_lineno)
        return false
# end of FileExist ################################################################

def fileSize(strFilePath):   ######################################################
    fileStatInfo = os.stat(strFilePath)
    myFileSize = fileStatInfo.st_size
    print_log("所取到的文件大小 = " + str(myFileSize))
    return myFileSize
# end of fileSize #################################################################

def ResizeImage(strSourcePicture_FullPath):  ######################################
    if not FileExist(strSourcePicture_FullPath):
        print_log("文件不存在,退出本轮的ResizeImage的处理", sys._getframe().f_lineno)
        return
    
    if fileSize(strSourcePicture_FullPath) <= realFileSizeInBytes_UpLimit:
        print_log("文件尺寸已经小于指定的上限,不必处理,本轮退出")
        return

    # 如果大小超过上限,则开始取数.
    img = Image.open(strSourcePicture_FullPath)
    (imgWidth, imgHeight) = img.size
    
    # 构造一个循环,让程序不断执行和优化.
    startRatio = 95
    currentRatio = startRatio
    
    startQuality = 95
    currentQuality = startQuality
    
    stepDownRatio = 5
    stepDownQuality = 5
    
    stepCount = 0
    stepMax = 30
    
    strSourceImage = strSourcePicture_FullPath
    strSourcImageBare = os.path.basename(os.path.realpath(strSourcePicture_FullPath))
    print_log("SourceImageBare = " + strSourcImageBare, sys._getframe().f_lineno)
    # 
    
    strRandomForTempImg = ''.join(choice(string.ascii_uppercase + string.digits) for i in range(8))
    
    strDestTempImage = os.path.join(TempFolder, strSourcImageBare + " " +  \
        time.strftime('%Y%m%d %H%M%S', time.localtime()) + " " + str(currentRatio) + "R " + \
        str(currentQuality) + "Q " + strRandomForTempImg + ".jpg")
        
    print_log("目标临时文件路径" + strDestTempImage, sys._getframe().f_lineno)
    while (stepCount <= stepMax):
        # 此处进入压缩状态.
        
        print_log("当前缩小比率 = " + str(currentRatio) + " 质量 = " + str(currentQuality), sys._getframe().f_lineno)
        newImg = img.resize((int(imgWidth * currentRatio / 100), int(imgHeight * currentRatio / 100)))
        newImg.save(strDestTempImage, "JPEG", quality=currentQuality )
        if stepCount % 2 == 0: 
            # 压缩尺寸
            currentRatio -= 5
        else:
            # 压缩质量            
            currentQuality -= 5
            
        # stepCount加一
        stepCount += 1
        
        if fileSize(strDestTempImage) <= realFileSizeInBytes_UpLimit:
            #作文件拷贝,然后退出循环        
            shutil.copy2(strDestTempImage, os.path.dirname(strSourcePicture_FullPath))
            break
        strDestTempImage = os.path.join(TempFolder, strSourcImageBare + " " +  \
            time.strftime('%Y%m%d %H%M%S', time.localtime()) + " " + str(currentRatio) + "R " + \
            str(currentQuality) + "Q " + strRandomForTempImg + ".jpg")
        print_log("目标临时文件路径" + strDestTempImage, sys._getframe().f_lineno)
        
        # end of loop
        
    print_log("上述图片的处理已经结束.转至下一轮次的图片", sys._getframe().f_lineno)
    
    
    
# end of ResizeImage  #############################################################


for iExistFile in range(len(existSourceFiles)):
    ResizeImage(existSourceFiles[iExistFile])

# 3. 读取源图片文件的参数：宽、高、质量。
# 4. 设定缩放比例的起始值、压缩质量的起始值。
# 5. 进行resize，并再做压缩quality
# 6. 验证目标图片文件的大小，符合要求，予以存放。

print_log('本程序的处理工作执行完毕')
# print("hello")
