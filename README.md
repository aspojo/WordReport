# WordReport
根据模版导出docx报表。
模版语法支持：变量表达式、foreach循环(支持无限级嵌套)、if、Merge(垂直方向上合并单元格)
具体用法，参考测试用例

Export the docx report according to the template. 
Template syntax support: variable expression, foreach loop (support infinite nesting), if, Merge (merge cells in the vertical direction).
for specific usage, refer to the test case
## 添加依赖
``` xml
<dependency>
    <groupId>top.logbug.wordreport</groupId>
    <artifactId>wordreport</artifactId>
    <version>1.0</version>
    <type>pom</type>
</dependency>
```
## 基本用法(Basic usage)：
``` java 
        Map<String, Object> context = new HashMap<>();
        Map<String, Object> relation = new HashMap<>();
        context.put("list", CollectionUtil.toList(relation));


        relation.put("upstreamDocName", "需求文档");
        relation.put("docName", "设计文档");
        relation.put("upstreamDocTitle", "需求文档");
        relation.put("docTitle", "设计文档");
        relation.put("upReqTitle", "需求");
        relation.put("upSectionTitle", "章节");
        relation.put("reqTitle", "标识");
        relation.put("sectionTitle", "章节");

        Map<String, Object> req = new HashMap<>();
        relation.put("reqList", CollectionUtil.toList(req, req));

        req.put("reqName", "设计1");
        req.put("section", "标题1");
        req.put("upReqName", "需求1");
        req.put("upSection", "标题1");

        WordReport.exportDocx("template.docx", context, "report.docx");
```
## 模版如下(Template)：
![image](https://user-images.githubusercontent.com/13494354/139300538-23c80399-57a4-4461-b043-93564f90c27f.png)
## 导出报表如下(Report)：
![image](https://user-images.githubusercontent.com/13494354/139300725-ce8dad90-c3f1-44ca-8241-8b0cdf30e4cc.png)
