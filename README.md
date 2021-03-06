# 应答器设置规则

## 一般规则

1. 应答器组内相邻应答器的距离为 5±0.5m
2. 发送线路参数的应答器组，车站正线组内应答器组距绝缘节距离不小于 30m。应答器组距绝缘节距离从靠近绝缘节的应答器计算。
3. CTCS-2 和 CTCS-3 级列控系统的应答器组内应答器数量不宜超过3个，发送线路参数的应答器组由两个及以上应答器构成。
4. 设置在车站的应答器组中的有源应答器应靠近信号机侧。
5. 两个应答器组链接距离不小于200m。

## 命名规则

1. 每个应答器 (组) 命名应以 B 开头，后加公里标或信号机名称，其中公里标参照区间通过信号机命名规则执行，即应答器名称以该应答器 (组) 所在位置坐标公里数和百米数组成，对于 km 后的单位采用四舍五入的方式计算,下行编号奇数，上行编号偶数。
2. 应答器名称应区分应答器组内的位置，分别在应答器名称后加 "-1","-2" 等表示组内第一个应答器和第二个应答器信息。

## 编号规则

1. 应答器编号应接 "大区号-分区号-车站号-应答器单元编号-组内编号" 格式填写。
2. 单元编号规则
    - 应答器单元编号以列车运行正方向或用途为参照，按正线贯通，从小到大的原则进行编号，下行为编号奇数 ，上行为偶数。
    （对于车站管辖范围内含区间的全部应答器组进行统一编号）
    - 单元编号由三位十进制表示，编号范围为 1-255。
3. 大区编号由三位十进制表示，编号范围为 1-127。
4. 分区号由一位十进制表示，编号范围为 1-7。
5. 车站编号规则
    - 车站编号由两位十进制表示，编号范围为 1-60，一个分区的车站数量一般不超过 50 个进行分配。
    - 接分区内车站的下行方向顺次进行车站编号。

### 思路

首先，我们要知道这个车站有多少个应答器组，假如车站有五个应答器组，然后他们的编号规则就是：'001','003','005'...

```python
for (i = 1;i < 11; i+2):
    if (i<10):
        num = '00' + str(i)
    else if(i<100):
        num = '0' + str(i)
    else:
        num = str(i)
    # 这个里 num 就是正确的单元编号
    print(num)
```

## 里程规则

1. 里程应填写应答器安装的实际线路运营里程 (格式为 KXXX + XXX)，精确到米，以靠近的信号机里程为参照点。

## 类型

1. 空心三角形 "△" 表示无源应答器。
2. 实心三角形 "▲" 表示有源应答器。

## 用途

1. CTCS-0 站应答器组 [cz-c0] 设置。
2. CTCS-0 车站向 CTCS-2 区域方向出站口 (含反向) 上下行各设置两个有源应答器组（由一个有源和两个及以上无源构成），向列车发送线路数据和临时限速信息。
3. CTCS-0/2等级转换应答器组（等级转换预告、执行）

   类型：应答器组包含两个无源应答器

   里程：
    1. 等级转换预告应答器组距等级转换执行应答器组的距离大于列车按等级转换点处线路最高允许速度运行5s的走行距离。（不大于160KM/h) 222m
    2. 等级执行应答器组设在距闭塞分区入口处30±0.5m处。（从靠近绝缘节的应答器算）处。
    3. cz-c02距zx要大于450m。
4. 定位应答器组：
    1. 仅用于定位的应答器组可为单个应答器
    2. 里程：
        - 在车站进站信号机 (含反向) 外方 250 ± 0.5m 外设置
5. 进站信号机应答器组：
    1. 类型： 进站信号机外方设置有源应答器组，包含两个无源
    2. 里程： 距进站信号机 30 ± 0.5m

## UML 数据建模属性

UML 模型如下

![应答器UML模型](./UML/Transponder.png)

属性描述：

| 属性       | 中文名 | 数据类型   |
| ---------- | ------ | ------ |
| B-Name     | 名称   | String |
| B-Num      | 编号   | String |
| B-Location | 里程   | String |
| B-Type     | 类型   | String |
| B_trueName | 正确名称 | String |
| B_trueNum  | 正确编号 | String |
| B_trueLocation | 正确里程 | String |
| B_trueType | 正确类型 | String |

方法描诉：

| 方法  | 作用 |
| ---- | ---- |
| verifyName | 验证名称 |
| verifyNum  | 验证编号 |

## 编程思路

首先我们安装操作 Excel 表的 xlrd,xlwt,xlutils 包。

然后先通过 xlrd 包，得到 Excel 中的信息，再通过循环遍历，得到每一行的数据，并存放到相应的数组

定义一个数组，用来存放应答器组，同时定义一个存放用途的数组，里面的数据与应答器组验证规则一一对应，后面代码在执行的时候，同时遍历这两个数组，来实现用途的验证

定义一些其他变量，实际上就是后面方法需要用到的数据，比如出站信号机位置等信息

定义一系列方法

- `getLocNum` 这个方法将里程转换为数字，需要传入一个里程参数
- `getUse` 这个方法通过一系列验证方法，例如等级转换等，得到这个车站正确的应答器组。放在一个数组里面返回
- `verifyExistence` 这个方法验证应答器是否缺失，如果缺失则返回真，否则，返回假
- `verifyLocation` 这个方法用来验证里程是否正确，它需要多个参数，包括 row(属于哪一行)，reference(参照点), B_Location(数据表中的里程), *args(任意类型参数，这个主要是对后面传递的参数不确定时使用)。该方法执行完毕会返回一个正确的里程供验证名称的时候使用
- `verifyName` 这个方法是用来验证名称是否正确，需要 row(属于哪一行), B_trueLocation(正确的里程), B_Name(数据表中的名称), use(用途), index(组内第几个)这几个参数。
- `verifyNum` 验证编号，需要 row(属于哪一行), B_Num(编号), index(组内第几个应答器)这几个参数，但是这个比较复杂，所有现在没有验证
- `verifyType` 验证类型，需要 row, use(用途，通过用途来判断这个组内是否有有源应答器), ponderType(类型), index(组内第几个应答器)这几个参数
- `_verifyType` 前面有个下划线，这是 `verifyType` 的私有方法，根据我们给的正确应答器类型来与传递进来的应答器类型判断，需要 row, ponderType(需要验证的应答器类型),trueTpye(正确的应答器类型) 这几个参数
- `verifyUse` 验证用途，需要 row, use(用途), trueUse(正确的用途)这几个参数
- `verify` 这个方法是用来标红的，上面的方法只要验证出了错误数据，都会调用这个方法，它需要 row(属于哪一行，也就是哪个应答器), col(哪一列，也就是验证的哪个属性), value(验证的这个属性的本来值), suggest(建议修改为正确值) 这几个属性

### 数据验证

首先执行 `getUse` 得到正确的应答器组，然后再通过 `while` 循环，循环遍历所有待验证数据，在循环中，通过实际的应答器组截取数据表中相应的应答器组，然后首先验证数据是否缺失，如果缺失，则执行下一个应答器，如果未缺失，则执行那些定义好的验证规则进行数据验证

验证完成之后调用 `xlwt` 中的 `save()` 方法导出一个新的，进行过数据验证的表

## 接下来的打算

1. 可能之前的思路出现了问题，我们可能需要重新整理一下思路，换个方向，不从应答器入手，而从验证规则入手建模。
2. 在接下来的重构代码中，如果从验证规则入手，就需要使用到策略模式，需要花时间研究。
3. 验证规则已经几乎全部投提取出来了，只有代码层面还没有实现，所以可以先写论文，边写论文边通过代码实现。

## 老师建议

- 根据规则弄一个正确的表出来，然后与旧表对比
- 编号问题，根据信号数据表，TCC边界与通过信号机之间的应答器都是属于左边车站