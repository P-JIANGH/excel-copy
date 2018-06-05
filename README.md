# 整理Excel工具（console）

## 将多个有共同识别特征，和相同内容格式的Excel文件copy到指定的Excel中

### 由于没有导出配置文件，使用前需要重写常量

* 代码如下

```c#
    #region 初期化定数の定義
    // 起始坐标
    const int START_ROW = 2;
    const int START_COLUME = 1;

    // 文件标识
    const string FILE_START_WITH = "";
    const string FILE_END_WITH = "";
    const string SHEET_NAME = "";

    // 目标文件标识
    const string NEW_FILE_NAME = "";
    const string NEW_SHEET_NAME = "";

    // 忽略文件夹
    public static readonly string[] IGNORE = { ".svn", ".git", "xx" };

    const int EOF = -1;
    #endregion
```

* 判断文件结束方式：连续三行某一列为空
* 目前代码不搜索Basepath下的指定文件
* 根据自己的需求，定制逻辑

```c#
    static void Main(string[] args)
    {
        // ...TODO
    }
```
