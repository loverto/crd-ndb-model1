// 导入大漠插件版本
const path = require("path");

const dm = require('@loverto/dm.dll')
const sleep = require('./sleep');
const fs = require('fs')
const readline = require('readline');

const dmExt = require('./dm.dll.ext')
const coreldraw = require('./coreldraw')
const common = require('./common')
const _ = require('lodash')

const keycode = require('keycode')

const log4js = require('log4js');
log4js.configure({
    appenders: {
        out: { type: 'stdout' },
        stb13: { type: 'file', filename: process.env.storePath+path.sep+'stb-13.log' } },
    categories: { default: { appenders: ['out','stb13'], level: 'debug' } }
});

const logger = log4js.getLogger('stb');

const {DB} = require('./database')

// 获取大漠插件的版本
logger.debug(dm.dll.ver())

let db = null;

const windowTitle = "CorelDRAW X6 (64 位)";
let devWidth = 1440
let devHeight = 900
let screenSize = dm.getScreenSize();
logger.debug(screenSize)
let winW = screenSize.width/devWidth;
let winH = screenSize.height/devHeight;

// dpi 界面坐标
let dpiUIPosition = [1021,629]

let dpiSwitchCoordinate = [871,529]

let dpiSwitchFlagCoordinate =  [869,548]
let dpiSwitchFlagColor =  "000000"
let dpiSwitchFlagSim =  0.8

// 需要修改的dpi的值
let dpiValue = 96

let fileSuff = ".cdr";

let modelFilePath = ""
let textFilePath = ""
let imageFilePath = ""
let exportModelFilePath = ""
let pch = ""

// 手提包 13面的旋转90度的 14个
let modelSpin = ["598,251","729,251","599,352","729,352","598,451","730,451","599,551","728,553","845,266","945,264","845,397","944,397","844,527","944,527"]
// 内胆包 模板6个位置的坐标
let modelInner = ["433,377","479,377","525,379","573,380","624,382","687,387"]
// 内胆包 图片6个位置的坐标
let imageInnerForCopyCoordinate = ["253,504","338,504","423,504","508,504","593,504"]
// 合并当前数组和imageInnerForCopyCoordinate
let imageInner = _.concat(["168,504"],imageInnerForCopyCoordinate)
// corel draw 调整元素高度的坐标
let adjustHeightUICoordinate = [181,103]
// 内胆包 图片的6个高度值
let imageInnerForWidthAndHeight = ["220,300","225,330","260,360","270,375","285,390","320,430"]

// 内胆包 复制次数
let innerCopyCount = 5
// 手提包 条数
const size = 1;

// 导出需要选择的坐标点
let exportCoordinate = ["408,272","583,435"]

// 箭头坐标
let arrowCoordinate = [12,136]


// 导入图片坐标
let importImagePositionCoordinate = ["149,481","428,540","618,536","808,533","990,531","223,538"]


// 左上角点击坐标
let leftClickCoordinate = importImagePositionCoordinate[0].split(",")

// 空白位置坐标
let clickWhite = [207,328]


// 模板坐标坐标
let modelCoordinate = importImagePositionCoordinate


// 宽和高位置
let widthAndHeightPosition = ["180,88","180,108"]
// 宽高参数
let widthHeightParam = ["389,247","357,255","347,245","325,229.5"]
// 参考行坐标
let refRowCoordinate =["396,495","526,498","667,503","800,498"]
// 批量复制左上右下坐标
let batchCopyCoordinate = ["120,453", "407,519"]

// 替换坐标
// 替换文本查找坐标, 替换文本替换坐标, 全部替换坐标,替换完成，替换关闭
let replaceCoordinate = common.ratioConversion(["614,396","610,426","907,455","748,484","966,363"],winW,winH)
// 替换需要查找的文本
let findText = "编号位置";
// 批次号增量标记
let pchIncreateFlag = -1


let refYValueOne = 103 * (winH)
let refYValueTwo = 76 * (winH)
let refYValueThree = 133 * (winH)
let refYValueFour = 64 * (winH)

let diffOne = 0;
let diffTwo = 0;

/*
* 按行读取文件内容
* 返回：字符串数组
* 参数：fReadName:文件名路径
*      callback:回调函数
* */
function readFileToArr(fReadName,callback){
    const fRead = fs.createReadStream(fReadName);
    const objReadline = readline.createInterface({
        input: fRead
    });
    const arr = new Array();
    objReadline.on('line',function (line) {
        arr.push(line);
        //console.log('line:'+ line);
    });
    objReadline.on('close',function () {
        // console.log(arr);
        callback(arr);
    });
}

/**
 * 异步读取文件目录下的所有符合条件的文件
 * @param dirPath 需要读取的文件目录
 * @param callback 回调函数
 */
function readDirToArrForCallback(dirPath,callback){
    let files = readDirToArr(dirPath,".JPG",true);
    logger.debug("dirPath",JSON.stringify(dirPath))
    logger.debug("files",JSON.stringify(files))
    callback(files)
    return files;
}
/**
 * 异步读取文件目录下的所有符合条件的文件
 * @param dirPath 需要读取的文件目录
 * @param suffix 后缀名称,默认为JPG
 * @param isIgnoreCase 是否忽略大小写
 * @param 返回所有的文件
 */
function readDirToArr(dirPath,suffix,isIgnoreCase){
    let files = fs.readdirSync(dirPath);
    const arr = new Array();
    for (let i =0; i<files.length; i++){
        let filename = files[i];
        let start = filename.lastIndexOf(".");
        // 文件名称不包含后缀
        let fileNameWithOutSuffix = filename.substring(0,start);
        // 从文件名称中截取后缀
        let orignSuffix = filename.substring(start,filename.length);
        logger.debug("orignSuffix",orignSuffix,"filename",filename)
        // 判断是否忽略大小写
        if (isIgnoreCase){
            if (orignSuffix.toLocaleLowerCase() == suffix.toLocaleLowerCase()){
                arr.push(fileNameWithOutSuffix);
                continue;
            }
        }else{
            if (orignSuffix == suffix){
                arr.push(fileNameWithOutSuffix);
                continue;
            }
        }
    }
    return arr;
}













/**
 * 主方法
 */
function main(configObject) {
    logger.debug("ndb model1 is starting up ")
    if (!configObject){
        let storePath = process.env.storePath;
        db = new DB(storePath);
        logger.debug("from db config")
        configObject  = db.get("configObject");
        logger.debug("config value " + JSON.stringify(configObject))
    }


    initConfig(configObject);
    // 如果没有找到窗口，则退出
    if (!coreldraw.findCorelDrawAndFullScreen(windowTitle)){
        logger.debug("corel draw window is not find")
        return;
    };

    //activeInput(windowTitle,"US")
    // return;
    logger.debug("corel draw eas")
    coreldraw.eas();
    logger.debug("corel draw start open model")
    coreldraw.openUModel(modelFilePath)
    sleep.msleep(1000)
    logger.debug("corel draw mouse is move arrow")
    // 设置为可移动
    coreldraw.moveAndClick(arrowCoordinate)
    // 调整文本文件路径为图像路径 textFilePath -> imageFilePath
    if (fs.existsSync(imageFilePath)){
        logger.debug("file is exists")
        //let readFileSync = fs.readFileSync(textFilePath);
        // 按行读取数据
        // let dataa = readDirToArr(imageFilePath);
        readDirToArrForCallback(imageFilePath,function (dataa) {
            logger.debug("data length"+dataa.length)

            let i =0;
            let j = 0
            // 先屏蔽该逻辑
            if (db.has('crd') && false){
                let crd = db.get('crd');
                logger.debug("crd from db config"+JSON.stringify(crd));
                // 当前执行的批次数
                i = crd.currentBatch;
                // 当前执行的条数
                j = crd.currentRow;
                // 获取缓存的批次号
                pch = crd.pch;
                // 获取缓存的批次号量
                pchIncreateFlag = crd.pchIncreateFlag;
            }
            let dataArr = _.chunk(dataa,size);
            logger.debug("dataArr ：" + dataArr.length);
            // 遍历按行读取的数据
            for (i; i<dataArr.length; i++){
                //db.set('crd',{currentBatch: i,pch:pch,pchIncreateFlag:pchIncreateFlag})
                let data = dataArr[i]
                for (let j = 0;j<data.length;j++){

                    let picfilename = data[j];
                    logger.debug("picfilename"+picfilename)
                    let picPath = common.getFilePathByFileName(imageFilePath,picfilename);
                    logger.debug("picPath"+picPath)
                    // 图片路径和模板路径都存在
                    if (fs.existsSync(picPath) && fs.existsSync(modelFilePath)){
                        handler(picPath,null,false,modelCoordinate,picfilename,j)
                    }
                    logger.debug("开始点击空白位置")
                    sleep.msleep(500)
                    coreldraw.moveAndClick(clickWhite)
                    sleep.msleep(500)

                    // 执行完数组中的值，就保存够13张则保存图片
                    if (j==data.length-1){
                        logger.debug("张数够了，开始保存")
                        logger.debug("开始获取序列号"+pch+pchIncreateFlag)
                        // 获取序列号
                        let result = common.getSequenceNumber(pch,pchIncreateFlag);
                        logger.debug("获取序列号后的结果"+JSON.stringify(result))
                        pch = result.pch;
                        pchIncreateFlag = result.pchIncreateFlag
                        let crd = {currentBatch:i,currentRow: j,pch:pch,pchIncreateFlag:pchIncreateFlag};
                        logger.debug("开始存储序列号到数据库中"+JSON.stringify(crd))
                        db.set('crd',crd)
                        logger.debug("开始替换编号")
                        // 替换编号
                        // coreldraw.findAndReplaceText(replaceCoordinate,findText,pch);
                        coreldraw.findAndReplaceText(replaceCoordinate,findText,picfilename);

                        sleep.msleep(200);
                        logger.debug("开始保存文件")
                        // let exportPathAbsout = exportModelFilePath + path.sep + pch + fileSuff;
                        let exportPathAbsout = exportModelFilePath + path.sep + picfilename + fileSuff;
                        logger.debug("开始另存为"+exportPathAbsout)
                        coreldraw.saveAsPath(exportPathAbsout);
                        sleep.msleep(3000)
                        logger.debug("保存完毕，开始关闭当前标签页")
                        coreldraw.closeModel();
                        coreldraw.closeModel();
                        coreldraw.closeModel();
                        sleep.msleep(500)
                        logger.debug("关闭完毕")
                        coreldraw.eas();
                        // 保存当前的序列号
                        if(db.has("configObject")){
                            configObject.pch = pch;
                            db.set("configObject",configObject)
                            logger.debug("把pch保存到数据库中"+JSON.stringify(configObject))
                        }

                    }

                }


                // // 执行完重置该行数据
                // j = 0;
                logger.debug("保存之后，判断是否需要打开新的模板")
                if (i<=dataArr.length-1){
                    logger.debug("执行完毕，开始保存，共执行"+i+"版");
                    sleep.msleep(500)
                    coreldraw.openUModel(modelFilePath);
                    sleep.msleep(200)
                    // 可移动坐标
                    coreldraw.moveAndClick(arrowCoordinate)
                }

            }

            // 最后执行完当前所有的图片后，编号自动更新一位，避免下次重命名
            // 获取序列号
            let result = common.getSequenceNumber(pch,pchIncreateFlag);
            pch = result.pch;
            pchIncreateFlag = result.pchIncreateFlag

            // 保存当前的序列号
            if(db.has("configObject")){
                configObject.pch = pch;
                db.set("configObject",configObject)
            }

            process.send({totalSize:dataa.length})

        })
    }

}

/**
 * 激活输入法
 * 该发放暂时不可用
 * @param windowTitle
 * @param input
 */
function activeInput(windowTitle,input) {
    const hwnd = dm.findWindow("", windowTitle);
    if (dmExt.checkInputMethod(hwnd, input) == 0) {
        dmExt.activeInputMethod(hwnd, input)
    }
}


/**
 * 核心处理业务方法
 * @param coreldrawHandlerFilePath 文件路径
 * @param model 模型
 * @param flag 标志位
 * @param coordinateArray 坐标点数组
 * @param filename 文件名称
 * @param number 当前张数
 */
function handler(coreldrawHandlerFilePath,model,flag,coordinateArray,filename,number) {
    logger.debug("corelDrawHandlerFilePath:"+coreldrawHandlerFilePath+"model:"+model
    + "flag:" + flag+"coordinateArray:"+coordinateArray+"filename:"+filename+
        "number:"+number
    )
    sleep.msleep(500)
    logger.debug("start import model")
    coreldraw.importUModel(coreldrawHandlerFilePath);
    sleep.msleep(500)
    logger.debug("start position"+JSON.stringify(leftClickCoordinate))
    dm.moveTo(leftClickCoordinate[0],leftClickCoordinate[1])
    sleep.msleep(200)
    dm.leftClick();
    sleep.msleep(2000)
    // 复制指定次数,并移动到对应位置
    logger.debug("start copy")
    coreldraw.copyObjectAndMoveCoordinate(imageInner[0].split(","),imageInnerForCopyCoordinate,clickWhite,innerCopyCount);
    logger.debug("start paste")
    // 点击空白坐标
    sleep.msleep(200)
    coreldraw.moveAndClick(clickWhite)
    sleep.msleep(200)

    // 复制完图片后,通过循环imageInner中的坐标元素,把鼠标移动到对应的坐标元素,然后在coreldraw界面上来一个一个的调整图片的高度
    for (let i = 0; i < imageInner.length; i++) {
        // 模型宽和高
        // let modelCurrentCoordinate = modelInner[i].split(",");
        // let width =[],height=[],widthTemp=0,heightTemp=0;
        // coreldraw.getObjectWidthAndHeight(modelCurrentCoordinate,width,height,widthTemp,heightTemp,widthAndHeightPosition);
        let coordinate = imageInner[i];
        logger.debug("start move"+JSON.stringify(coordinate))
        coreldraw.moveAndClick(coordinate.split(","))
        sleep.msleep(200)
        // logger.debug("start adjust"+JSON.stringify(width)+"ImageInnerForHeight:"+JSON.stringify(heightTemp)+"widthAndHeightPosition:"+JSON.stringify(widthAndHeightPosition)+"height:"+JSON.stringify(height)+"widthTemp:"+JSON.stringify(widthTemp)+"heightTemp:"+JSON.stringify(heightTemp))
        let widthAndHeight = imageInnerForWidthAndHeight[i].split(",");
        // 调整图片的高度
        // coreldraw.moveSpinNumber(widthAndHeightPosition[0].split(","),widthAndHeight[0]);
        // sleep.msleep(200)
        coreldraw.moveSpinNumber(widthAndHeightPosition[1].split(","),widthAndHeight[1]);
        sleep.msleep(200)
    }
    // 点击空白坐标
    sleep.msleep(200)
    coreldraw.moveAndClick(clickWhite)
    sleep.msleep(200)
    // 开始把图片移动到模板位置modelInner中
    // 1. 循环modelInner中的元素
    // 2. 先按住shift键
    // 2. 然后选中第一张图和第一个模板位置
    // 3. 然后按C+E键,然后按住shift键,选中第二张图和第二个模板位置.以此类推
    for (let i = 0;i<imageInner.length;i++){
        let moveCoordinate = imageInner[i].split(",");
        let endCoordinate = modelInner[i].split(",");
        logger.debug("start move"+JSON.stringify(moveCoordinate)+"endCoordinate:"+JSON.stringify(endCoordinate))
        coreldraw.drag(moveCoordinate,endCoordinate)
        // // 通过快捷键方式对齐，比通过坐标对其更精准
        // dm.keyDown(keycode("shift"))
        // coreldraw.moveAndClick(moveCoordinate);
        // sleep.msleep(200);
        // coreldraw.moveAndClick(endCoordinate);
        // sleep.msleep(200);
        // dm.keyUp(keycode("shift"));
        // sleep.msleep(500);
        // dm.keyPress(keycode("shift"));
        // sleep.msleep(500);
        // dm.keyPress(keycode("c"));
        // dm.keyPress(keycode("e"));
        // 点击空白坐标
        sleep.msleep(200)
        coreldraw.moveAndClick(clickWhite)
        sleep.msleep(200)
    }
}

/**
 * 初始化配置文件
 * @param configObject 配置对象
 */
function initConfig(configObject) {
    modelFilePath = configObject.modelFilePath;
    imageFilePath = configObject.imageFilePath;
    exportModelFilePath = configObject.exportModelFilePath;
    textFilePath = configObject.textFilePath;
    pch = configObject.pch;
}

exports.main = main


main();
