# LambdaIO

使用Lambda表达式便捷的进行导入导出操作的.NET类库，支持DataTable、CSV、Excel(NPOI)

## 缘起

平常开发中难免遇到集合与外部数据导入导出的需求，如何表达导入导出字段的对应关系是个问题，传统做法是使用Attribute在属性上标记，这存在三大问题：
1. .NET新增的动态类型和匿名类型无法设置Attribute，如果为每个场景都创建数据结构显得浪费；
2. Attribute不够灵活，比如无法表达方法、表达式、顺序等；
3. Attribute不好约束类型，导致不可避免的拆箱装箱，影响性能。

在下在2年前的隔壁项目[DotNetUtility](https://github.com/starchen88/DotNetUtility)中已经实现了通过委托来表达导入导出关系的逻辑：
```
var dt = OfficeHelper.ToDatatable(CreateTestData1List(), new Dictionary<string, Expression<Func<TestData1, object>>> {
    { "Id",it=>it.Id },
    { "Name",it=>it.Name },
    { "Time",it=>it.Time}
});
```
但是这里的委托都是object类型，还是无法避免拆箱装箱问题，而且使用略繁琐。

开这个新的项目主要是重构之前的设计，一是**添加对于表达式树的支持！**，二是添加对于CSV格式的支持。

## 如何使用

快马加鞭更新中

## 感谢

造轮子不易！写bug也不易！请多支持！分享、加星、反馈、参与话题、贡献代码、赞助……都是极好的！