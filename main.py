from mylogger import logger
from xl2dict import XlToDict
import os
import glob

class DoctorPhotoClean():
    def __init__(self, excelFilePath, inputDir, outputDir):
        self.excelFilePath = excelFilePath
        self.inputDir = inputDir
        self.outputDir = outputDir
        if not os.path.exists(self.outputDir):
            logger.info("{} does not exist, create this output directory", self.outputDir)
            os.mkdir(self.outputDir)
        self.parseData()
    
    def parseData(self):
        myxlobject= XlToDict()
        self.dictData = myxlobject.convert_sheet_to_dict(file_path=self.excelFilePath)

    def sourceFiles(self):
            sourceFiles = []
            for file in glob.glob(self.inputDir + os.path.sep + "*.jpg"):
                sourceFiles.append(file) 
            return sourceFiles

    def doTransform(self):
        sourceFiles = self.sourceFiles()
        logger.info("transform {} source files", len(sourceFiles))
        for dataEntry in self.dictData:
            logger.debug("process 学号: {}/报名号: {}", int(dataEntry['学号']), int(dataEntry['报名号']))
            sourcePath = os.path.join(self.inputDir, str(int(dataEntry['报名号'])) + ".jpg")
            targetPath = os.path.join(self.outputDir, str(int(dataEntry['学号'])) + ".jpg")
            if not os.path.exists(sourcePath):
                logger.warn("{} does not exit, skip it", sourcePath)
            else:
                logger.debug("rename {} to {}", sourcePath, targetPath)
                os.rename(sourcePath, targetPath)

class GraduatePhotoClean(DoctorPhotoClean):
    def __init__(self, excelFilePath, inputDir, outputDir):
        DoctorPhotoClean.__init__(self, excelFilePath, inputDir, outputDir)

    def doTransform(self):
        sourceFiles = self.sourceFiles()
        logger.info("transform {} source files", len(sourceFiles))
        for dataEntry in self.dictData:
            logger.debug("process 学号: {}/报名号: {}", int(dataEntry['学号']), dataEntry['报名号'])
            sourcePath1 = os.path.join(self.inputDir, str(dataEntry['证件号码']) + ".jpg")
            sourcePath2 = os.path.join(self.inputDir, str(int(dataEntry['考生编号'])) + ".jpg")
            sourcePath3 = os.path.join(self.inputDir, '10673_18_' + str(dataEntry['报名号']) + ".jpg")
            targetPath = os.path.join(self.outputDir, str(int(dataEntry['学号'])) + ".jpg")
            if os.path.exists(sourcePath1):
                logger.debug("rename from (证件号码) {} to {}", sourcePath1, targetPath)
                os.rename(sourcePath1, targetPath)
            elif os.path.exists(sourcePath2):
                logger.debug("rename from (考生编号) {} to {}", sourcePath2, targetPath)
                os.rename(sourcePath2, targetPath)
            elif os.path.exists(sourcePath3):
                logger.debug("rename from (10673_18_报名号) {} to {}", sourcePath3, targetPath)
                os.rename(sourcePath3, targetPath)
            else:
                logger.warn("source file for {} does not exit, skip it", int(dataEntry['学号']))

doctorPhotoClean = DoctorPhotoClean(r'D:\code\ynu\clean-photo\2018-博士.xlsx', 
    r"C:\Users\Liu.D.H\Documents\photo\硕博照片\2018博士照片",
    r"C:\Users\Liu.D.H\Documents\photo\硕博照片\2018博士照片-transform")
doctorPhotoClean.doTransform()

graduatePhotoClean = GraduatePhotoClean(r'D:\code\ynu\clean-photo\2018-硕士.xlsx', 
    r"C:\Users\Liu.D.H\Documents\photo\硕博照片\2018硕士照片",
    r"C:\Users\Liu.D.H\Documents\photo\硕博照片\2018硕士照片-transform")
graduatePhotoClean.doTransform()
