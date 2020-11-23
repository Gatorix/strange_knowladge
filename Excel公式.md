# Excel公式

## 常量

## 运算符

### 算数运算符

- + _ * / % ^

### 比较运算符

- = > < <= >= <> !=

### 引用运算符

- 区域

	- :

- 联合

	- ,

- 交叉

	- 空格

### 文本运算符

- &

## 引用

### 引用类型

- 绝对引用
- 相对引用
- 混合引用

### 跨工作表引用

- =Sheet1!A2

### 跨工作簿引用

- =[Book1]Sheet1!A2

## 函数

### 逻辑函数

- AND

	- 所有参数的计算结果为 TRUE 时，AND 函数返回 TRUE；只要有一个参数的计算结果为 FALSE，即返回 FALSE。
	- AND(logical1, [logical2], ...)

		- 参数是计算结果可为True或False的条件
		- logical1必填
		- logical2…参数选填，最多255个条件

- OR

	- 如果 OR 函数的任意参数计算为 TRUE，则其返回 TRUE；如果其所有参数均计算机为 FALSE，则返回 FALSE。
	- OR(logical1, [logical2], ...)

		- 同AND

- NOT

	- 对其参数的逻辑求反
	- NOT(logical)

		- =NOT(TRUE)返回FALSE
		- =NOT(FALSE)返回TRUE

- IF

	- 使用逻辑函数 IF 函数时，如果条件为真，该函数将返回一个值；如果条件为假，函数将返回另一个值。
	- IF(logical_test, value_if_true, [value_if_false])

		- logical_test   （必需）要测试的条件。
		- value_if_true   （必需）logical_test 的结果为 TRUE 时，您希望返回的值。
		- value_if_false   （可选）logical_test 的结果为 FALSE 时，您希望返回的值。该参数不录入时直接返回FALSE

	- 延伸

		- IFERROR

			- 可以使用 IFERROR 函数捕获和处理公式中的错误。 如果公式的计算结果为错误值，则 IFERROR 返回您指定的值;否则，它将返回公式的结果。
			- IFERROR(value, value_if_error)

				- value（必需）检查是否存在错误的参数。
				- value_if_error    必需。 公式计算错误时返回的值。 计算以下错误类型：

					-  #N/A

						- 当数值对函数或公式不可用时，出现该错误 

					- #VALUE！

						- 当在公式或函数中使用的参数或操作数类型错误时，出现该错误 

					- #REF！

						- 当单元格引用无效时，出现该错误 

					- #DIV/0！

						- 当数字除以 0 时，出现该错误

					- #NUM！

						- 如果公式或函数中使用了无效的数值，出现该错误    

					- #NAME？

						- 当 Excel 无法识别公式中的文本时，出现该错误

					-  #NULL！

						- 如果指定两个并不相交的区域的交点，出现该错误

		- IFS（2019）

			- IFS 函数检查是否满足一个或多个条件，且返回符合第一个 TRUE 条件的值。 IFS 可以取代多个嵌套 IF 语句，并且有多个条件时更方便阅读。
			- IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], [logical_test3, value_if_true3],…)

				- logical_test1（必需）

					- 计算结果为 TRUE 或 FALSE 的条件。

				- value_if_true1（必需）

					- 当 logical_test1 的计算结果为 TRUE 时要返回结果。 可以为空。

				- 可选条件最多127个

					- logical_test2…logical_test127（可选）
					- value_if_true2…value_if_true127（可选）

### 数学、统计函数

- AVERAGE

	- 求出所有参数的算术平均值。如果某个单元格是空的或包含文本，它将不用于计算平均数。如果单元格数值为0，将参于计算平均数
	- AVERAGE( number, number2,……)

		- 其中 number1为必需的，后续值是可选的，是需要计算平均值的1到255个数值参数。

	- 延伸

		- AVERAGEA

			- 计算参数列表中数值的平均值（算术平均值）。
			- AVERAGEA(value1, [value2], ...)

				- Value1, value2, ...    Value1 是必需的，后续值是可选的。 需要计算平均值的 1 到 255 个单元格、单元格区域或值。

		- AVERAGEIF

			- 返回某个区域内满足给定条件的所有单元格的平均值（算术平均值）。
			- AVERAGEIF(range, criteria, [average_range])

				- Range    必需。 要计算平均值的一个或多个单元格，其中包含数字或包含数字的名称、数组或引用。
				- Criteria    必需。 形式为数字、表达式、单元格引用或文本的条件，用来定义将计算平均值的单元格。 例如，条件可以表示为 32、"32"、">32"、"苹果" 或 B4。
				- Average_range    可选。 计算平均值的实际单元格组。 如果省略，则使用 range。

		- AVERAGEIFS（2019）

			- 返回满足多个条件的所有单元格的平均值（算术平均值）。
			- AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)

				- Average_range    必需。 要计算平均值的一个或多个单元格，其中包含数字或包含数字的名称、数组或引用。
				- Criteria_range1、criteria_range2 等    Criteria_range1 是必需的，后续 criteria_range 是可选的。 在其中计算关联条件的 1 至 127 个区域。
				- Criteria1、criteria2 等    Criteria1 是必需的，后续 criteria 是可选的。 形式为数字、表达式、单元格引用或文本的 1 至 127 个条件，用来定义将计算平均值的单元格。 例如，条件可以表示为 32、"32"、">32"、"苹果" 或 B4。

- MIN

	- 返回一组值中的最小值。
	- MIN(number1, [number2], ...)

		- number1, number2, ...    number1 是可选的，后续数字是可选的。 要从中查找最小值的 1 到 255 个数字。

- MAX

	- 返回一组值中的最大值。
	- MAX(number1, [number2], ...)

		- number1, number2, ...    Number1 是必需的，后续数字是可选的。 要从中查找最大值的 1 到 255 个数字。

- COUNT

	- COUNT 函数计算包含数字的单元格个数以及参数列表中数字的个数。 使用 COUNT 函数获取区域中或一组数字中的数字字段中条目的个数。
	- COUNT(value1, [value2], ...)

		- value1    必需。 要计算其中数字的个数的第一项、单元格引用或区域。
		- value2, ...    可选。 要计算其中数字的个数的其他项、单元格引用或区域，最多可包含 255 个。

	- 延伸

		- COUNTIF

			- 用于统计满足某个条件的单元格的数量
			- COUNTIF(range, criteria)

				- range   （必需）

					- 要进行计数的单元格组。 区域可以包括数字、数组、命名区域或包含数字的引用。 空白和文本值将被忽略。

				- criteria   （必需）

					- 用于决定要统计哪些单元格的数量的数字、表达式、单元格引用或文本字符串。
					- 例如，可以使用 32 之类数字，“>32”之类比较，B4 之类单元格，或“苹果”之类单词。
					- COUNTIF 仅使用一个条件。 如果要使用多个条件，请使用 COUNTIFS。

		- COUNTIFS（2019）

			- COUNTIFS 函数将条件应用于跨多个区域的单元格，然后统计满足所有条件的次数。
			- COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2],…)

				- criteria_range1    必需。 在其中计算关联条件的第一个区域。
				- criteria1    必需。 条件的形式为数字、表达式、单元格引用或文本，它定义了要计数的单元格范围。 例如，条件可以表示为 32、">32"、B4、"apples"或 "32"。
				- criteria_range2, criteria2, ...    可选。 附加的区域及其关联条件。 最多允许 127 个区域/条件对。

- ABS

	- 返回数字的绝对值。
	- ABS(number)

		- Number    必需。 需要计算其绝对值的实数。

- ROUND

	- ROUND 函数将数字四舍五入到指定的位数。
	- ROUND(number, num_digits)

		- number    必需。 要四舍五入的数字。
		- num_digits    必需。 要进行四舍五入运算的位数。

			- 如果 num_digits 大于 0，则将数字四舍五入到指定的小数位数。
			- 如果 num_digits 等于 0，则将数字四舍五入到最接近的整数。
			- 如果 num_digits 小于 0，则将数字四舍五入到小数点左边的相应位数。

- SUM

	- SUM函数将为值求和。Alt+=
	- SUM(number1,[number2],...)

		- number1   （必需）

			- 要相加的第一个数字。 该数字可以是 4 之类的数字，B6 之类的单元格引用或 B2:B8 之类的单元格范围。

		- number 2-255   （可选）

			- 这是要相加的第二个数字。 可以按照这种方式最多指定 255 个数字。

	- 延伸

		- SUMIF

			- 可以使用 SUMIF 函数对 范围 中符合指定条件的值求和。 
			- SUMIF(range, criteria, [sum_range])

				- range   必需。 要按条件计算的单元格区域。 每个区域中的单元格都必须是数字，或者是包含数字的名称、数组或引用。 空白和文本值将被忽略。 
				- criteria   必需。 定义哪些单元格将被添加的数字、表达式、单元格引用、文本或函数形式的条件。 可以包含通配符字符-问号（？）匹配任意单个字符，星号（*）匹配任何字符序列。 如果要查找实际的问号或星号，请在该字符前键入波形符 (~)。
				- sum_range   可选。 要添加的实际单元格（如果要添加的单元格不在range参数中指定的单元格）。 如果省略了sum_range参数，则 Excel 将添加在range参数中指定的单元格（在应用条件的相同单元格）。

		- SUMIFS（2019）

			- 用于计算其满足多个条件的全部参数的总量。
			- SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)

				- Sum_range   （必需）

					- 要求和的单元格区域。

				- Criteria_range1   （必需）

					- 使用 Criteria1 测试的区域。

				- Criteria1   （必需）

					- 定义将计算 Criteria_range1 中的哪些单元格的和的条件。 例如，可以将条件输入为 32、">32"、B4、"苹果" 或 "32"。

				- Criteria_range2, criteria2, …   

					- 附加的区域及其关联条件。 最多可以输入 127 个区域/条件对。

### 文本函数

- FIND

	- 用于在第二个文本串中定位第一个文本串，并返回第一个文本串的起始位置的值，该值从第二个文本串的第一个字符算起。

		- 无论默认语言设置如何，函数 FIND 始终将每个字符（不管是单字节还是双字节）按 1 计数。

	- FIND(find_text, within_text, [start_num])

		- find_text    必需。 要查找的文本。
		- within_text    必需。 包含要查找文本的文本。
		- start_num    可选。 指定开始进行查找的字符。 within_text 中的首字符是编号为 1 的字符。 如果省略 start_num，则假定其值为 1。

- EXACT

	- 比较两个文本字符串，如果它们完全相同，则返回 TRUE，否则返回 FALSE。 函数 EXACT 区分大小写，但忽略格式上的差异。 使用 EXACT 可以检验在文档中输入的文本。
	- EXACT(text1, text2)

		- Text1    必需。 第一个文本字符串。
		- text2    必需。 第二个文本字符串。

- TEXT

	- TEXT 函数可通过格式代码向数字应用格式，进而更改数字的显示方式。
	- TEXT(value, format_text)

		- value

			- 要转换为文本的数值。

		- format_text

			- 一个文本字符串，定义要应用于所提供值的格式。

				- "$#,##0.00"

					- 货币格式

				- "MM/DD/YYYY"

					- 日期格式

				- "0000000"

					- 补零

				- "0.00E+00"

					- 科学计数法

- LEN

	- LEN 返回文本字符串中的字符个数。
	- LEN(text)

		- Text    必需。 要查找其长度的文本。 空格将作为字符进行计数。

- LEFT

	- LEFT 从文本字符串的第一个字符开始返回指定个数的字符。
	- LEFT(text, [num_chars])

		- Text    必需。 包含要提取的字符的文本字符串。
		- num_chars    可选。 指定要由 LEFT 提取的字符的数量。

			- Num_chars 必须大于或等于零。
			- 如果 num_chars 大于文本长度，则 LEFT 返回全部文本。
			- 如果省略 num_chars，则假定其值为 1。

- MID

	- MID 返回文本字符串中从指定位置开始的特定数目的字符
	- MID(text, start_num, num_chars)

		- text    必需。 包含要提取字符的文本字符串。
		- start_num    必需。 文本中要提取的第一个字符的位置。 文本中第一个字符的 start_num 为 1，以此类推。

			- 如果 start_num 大于文本长度，则 MID/MIDB 返回 "" （空文本）。
			- 如果 start_num 小于文本的长度，但 start_num 加 num_chars 超过文本的长度，则 MID/MIDB 返回文本结尾的字符。
			- 如果 start_num 小于1，MID/MIDB 将返回 #VALUE！ 。

		- num_chars    对 MID 是必需的。 指定希望 MID 从文本中返回字符的个数。

			- 如果 num_chars 为负值，MID 将返回 #VALUE！ 。

- RIGHT

	- 根据所指定的字符数返回文本字符串中最后一个或多个字符。
	- RIGHT(text,[num_chars])

		- text    必需。 包含要提取字符的文本字符串。
		- num_chars    可选。 指定希望 RIGHT 提取的字符数。

			- Num_chars 必须大于或等于零。
			- 如果 num_chars 大于文本长度，则 RIGHT 返回所有文本。
			- 如果省略 num_chars，则假定其值为 1。

- REPLACE

	- 根据指定的字符数，REPLACE 将部分文本字符串替换为不同的文本字符串。
	- REPLACE(old_text, start_num, num_chars, new_text)

		- old_text    必需。 要替换其部分字符的文本。
		- start_num    必需。 old_text 中要替换为 new_text 的字符位置。
		- num_chars    必需。 old_text 中希望 REPLACE 使用 new_text 来进行替换的字符数。
		- Num_bytes    必需。 old_text 中希望 REPLACEB 使用 new_text 来进行替换的字节数。
		- new_text    必需。 将替换 old_text 中字符的文本。

- TRIM

	- 除了单词之间的单个空格之外，移除文本中的所有空格。 对于从另一个可能含有不规则间距的应用程序收到的文本，可以使用 TRIM。
	- TRIM(text)

		- Text    必需。 要从中移除空格的文本。

### 日期时间函数

- DATE

	- DATE 函数返回表示特定日期的连续序列号。
	- DATE(year,month,day)

		- Year   ：必需。year 参数的值可以包含一到四位数字。Excel 将根据计算机正在使用的日期系统来解释 year 参数。默认情况下，Microsoft Excel for Windows 使用的是 1900 日期系统，这表示第一个日期为 1900 年 1 月 1 日。

			- 如果 year 介于 0（零）到 1899 之间（包含这两个值），则 Excel 会将该值与 1900 相加来计算年份。例如，DATE(108,1,2) 返回 2008 年 1 月 2 日 (1900+108)。
			- 如果 year 介于 1900 到 9999 之间（包含这两个值），则 Excel 将使用该数值作为年份。例如，DATE(2008,1,2) 将返回 2008 年 1 月 2 日。
			- 如果 year 小于 0 或大于等于 10000，则 Excel 返回 错误值 #NUM!。

		- Month    必需。一个正整数或负整数，表示一年中从 1 月至 12 月（一月到十二月）的各个月。

			- 如果 month 大于 12，则 month 会从指定年份的第一个月开始加上该月份数。例如，DATE(2008,14,2) 返回表示 2009 年 2 月 2 日的序列数。
			- 如果 month 小于 1，则 month 会从指定年份的第一个月开始减去该月份数，然后再加上 1 个月。例如，DATE(2008,-3,2) 返回表示 2007 年 9 月 2 日的序列号。

		- Day    必需。一个正整数或负整数，表示一月中从 1 日到 31 日的各天。

			- 如果 day 大于指定月中的天数，则 day 会从该月的第一天开始加上该天数。例如，DATE(2008,1,35) 返回表示 2008 年 2 月 4 日的序列数。
			- 如果 day 小于 1，则 day 从指定月份的第一天开始减去该天数，然后再加上 1 天。例如，DATE(2008,1,-15) 返回表示 2007 年 12 月 16 日的序列号。

- YEAR

	- 返回对应于某个日期的年份。 Year 作为 1900 - 9999 之间的整数返回。
	- YEAR(serial_number)

		- Serial_number    必需。 要查找的年份的日期。 应使用 DATE 函数输入日期，或者将日期作为其他公式或函数的结果输入。 例如，使用函数 DATE(2008,5,23) 输入 2008 年 5 月 23 日。 如果日期以文本形式输入，则会出现问题。

- MONTH

	- 返回日期（以序列数表示）中的月份。 月份是介于 1（一月）到 12（十二月）之间的整数。
	- MONTH(serial_number)

		- Serial_number    必需。 您尝试查找的月份的日期。 应使用 DATE 函数输入日期，或者将日期作为其他公式或函数的结果输入。 例如，使用函数 DATE(2008,5,23) 输入 2008 年 5 月 23 日。 如果日期以文本形式输入，则会出现问题。

- DAY

	- 返回以序列数表示的某日期的天数。 天数是介于 1 到 31 之间的整数。
	- DAY(serial_number)

		- Serial_number    必需。 你尝试查找的日期的日期。 应使用 DATE 函数输入日期，或者将日期作为其他公式或函数的结果输入。 例如，使用函数 DATE(2008,5,23) 输入 2008 年 5 月 23 日。 如果日期以文本形式输入，则会出现问题。

- DAYS（2013）

	- 返回两个日期之间的天数。
	- DAYS(end_date, start_date)

		- End_date    必需。 Start_date 和 End_date 是用于计算期间天数的起止日期。
		- Start_date    必需。Start_date 和 End_date 是用于计算期间天数的起止日期。

- NOW

	- 返回当前日期和时间的序列号。
	- Now()

		- NOW 函数语法没有参数。

### 查找引用

- HLOOKUP

	- 在表格的首行或数值数组中搜索值，然后返回表格或数组中指定行的所在列中的值。 当比较值位于数据表格的首行时，如果要向下查看指定的行数，则可使用 HLOOKUP。 当比较值位于所需查找的数据的左边一列时，则可使用 VLOOKUP。
	- HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])

		- Lookup_value    必需。 要在表格的第一行中查找的值。 Lookup_value 可以是数值、引用或文本字符串。
		- Table_array    必需。 在其中查找数据的信息表。 使用对区域或区域名称的引用。

			- Table_array 的第一行的数值可以为文本、数字或逻辑值。
			- 如果 range_lookup 为 TRUE，则 table_array 的第一行的数值必须按升序排列：...-2、-1、0、1、2、...、A-Z、FALSE、TRUE；否则，HLOOKUP 将不能给出正确的数值。 如果 range_lookup 为 FALSE，则 table_array 不必进行排序。
			- 文本不区分大小写。
			- 将数值从左到右按升序排序。

		- Row_index_num    必需。 Table_array 中将返回匹配值的行号。 Row_index_num 1 返回 table_array 中的第一行值，row_index_num 2 返回 table_array 等中的第二行值。 如果 row_index_num 小于1，则 HLOOKUP 返回 #VALUE！ 错误值;如果 row_index_num 大于 table_array 中的行数，则 HLOOKUP 返回 #REF！ 。
		- Range_lookup    可选。 一个逻辑值，指定希望 HLOOKUP 查找精确匹配值还是近似匹配值。 如果为 TRUE 或省略，则返回近似匹配值。 换言之，如果找不到精确匹配值，则返回小于 lookup_value 的最大值。 如果为 False，则 HLOOKUP 将查找精确匹配值。 如果找不到精确匹配值，则返回错误值 #N/A。

- VLOOKUP

	- 当需要在表格或区域中按行查找项目时，请使用 VLOOKUP。 
	- VLOOKUP (lookup_value, table_array, col_index_num, [range_lookup])

		- lookup_value   （必需参数）

			- 要查找的值。 要查找的值必须位于您在table_array参数中指定的单元格区域的第一列中。

		- Table_array   （必需参数）

			- 要查找的位置。VLOOKUP 在其中搜索 lookup_value 和返回值的单元格区域。 你可以使用命名区域或表，并且可以在参数中使用名称，而不是单元格引用。 

		- col_index_num   （必需参数）

			- 包含要返回的值的区域中的列号（从1开始的table_array的最左侧列）。

		- range_lookup   （可选参数）

			- 一个逻辑值，该值指定希望 VLOOKUP 查找近似匹配还是精确匹配：

				- 近似匹配-1/TRUE假设表中的第一列按数值或字母顺序排序，然后将搜索最接近的值。 这是未指定值时的默认方法。 例如，= VLOOKUP （90，A1： B100，2，TRUE）。
				- 完全匹配-0/FALSE将搜索第一列中的确切值。 例如，= VLOOKUP （"Smith"，A1： B100，2，FALSE）。

- LOOKUP

	- 官方用法

		- 向量形式

			- LOOKUP 的向量形式在单行区域或单列区域（称为“向量”）中查找值，然后返回第二个单行区域或单列区域中相同位置的值。
			- LOOKUP(lookup_value, lookup_vector, [result_vector])

				- lookup_value    必需。 LOOKUP 在第一个向量中搜索的值。 Lookup_value 可以是数字、文本、逻辑值、名称或对值的引用。
				- lookup_vector    必需。 只包含一行或一列的区域。 lookup_vector 中的值可以是文本、数字或逻辑值。

		- 数组形式

			-  强烈建议使用 VLOOKUP 或 HLOOKUP，不要使用数组形式。
			- LOOKUP 的数组形式在数组的第一行或第一列中查找指定的值，并返回数组最后一行或最后一列中同一位置的值。 当要匹配的值位于数组的第一行或第一列中时，请使用 LOOKUP 的这种形式。
			- LOOKUP(lookup_value, array)

				- lookup_value    必需。 LOOKUP 在数组中搜索的值。 lookup_value 参数可以是数字、文本、逻辑值、名称或对值的引用。

					- 如果 LOOKUP 找不到 lookup_value 的值，它会使用数组中小于或等于 lookup_value 的最大值。
					- 如果 lookup_value 的值小于第一行或第一列中的最小值（取决于数组维度），LOOKUP 会返回 #N/A 错误值。

				- array    必需。 包含要与 lookup_value 进行比较的文本、数字或逻辑值的单元格区域。

					- 如果数组包含宽度比高度大的区域（列数多于行数）LOOKUP 会在第一行中搜索 lookup_value 的值。
					- 如果数组是正方的或者高度大于宽度（行数多于列数），LOOKUP 会在第一列中进行搜索。
					- 使用 HLOOKUP 和 VLOOKUP 函数，您可以通过索引以向下或遍历的方式搜索，但是 LOOKUP 始终选择行或列中的最后一个值。

	- 民间用法

		- 逆向查询

			- =LOOKUP(1,0/(B11:B23=A27),A11:A23)

		- 区间查询

			- =LOOKUP(D11,{0,30,50,60,80;"一般","良好","较好","优秀","能手"})

		- 查询列中的最后一个值

			- =LOOKUP("座",A11:A23)

				- 查文本

			- =LOOKUP(9E307,D11:D23)

				- 查数字

			- =LOOKUP(1,0/(B11:B23<>""),B11:B23)

				- 查文本数字组合

		- 根据简称查全称

			- =IFERROR(LOOKUP(1,0/FIND(A37,$A$11:$A$23),$A$11:$A$23),"无")

- INDEX

	- INDEX 函数返回表格或区域中的值或值的引用。
	- 数组形式

		- 返回表或数组中元素的值，由行号和列号索引选择。

			- 当函数 INDEX 的第一个参数为数组常量时，使用数组形式。

		- INDEX(array, row_num, [column_num])

			- array    必需。 单元格区域或数组常量。

				- 如果数组只包含一行或一列，则相应的 row_num 或 column_num 参数是可选的。
				- 如果数组具有多行和多列，并且仅使用 row_num 或 column_num，则 INDEX 返回数组中整个行或列的数组。

			- row_num    必需，除非存在 column_num。 选择数组中的某行，函数从该行返回数值。 如果省略 row_num，则需要 column_num。
			- column_num    可选。 选择数组中的某列，函数从该列返回数值。 如果省略 column_num，则需要 row_num。

	- 引用形式

		- 返回指定的行与列交叉处的单元格引用。 如果引用由非相邻的选项组成，则可以选择要查找的选择内容。
		- INDEX(reference, row_num, [column_num], [area_num])

			- reference    必需。 对一个或多个单元格区域的引用。

				- 如果要为引用输入非相邻区域，请将引用括在括号中。
				- 如果引用中的每个区域仅包含一行或一列，则 "row_num" 或 "column_num" 参数分别是可选的。 例如，对于单行的引用，可以使用函数 INDEX(reference,,column_num)。

			- row_num    必需。 引用中某行的行号，函数从该行返回一个引用。
			- column_num    可选。 引用中某列的列标，函数从该列返回一个引用。
			- area_num    可选。 选择一个引用区域，从该区域中返回 row_num 和 column_num 的交集。 选择或输入的第一个区域的编号为1，第二个区域为2，依此类推。 如果省略 area_num，则 INDEX 使用区域1。  此处列出的区域必须位于一个工作表上。  如果你指定的区域不在同一工作表上，它将导致 #VALUE！ 错误。  如果需要使用彼此位于不同工作表上的区域，建议使用 INDEX 函数的数组形式，并使用另一个函数计算构成数组的区域。  例如，可以使用 CHOOSE 函数计算将使用的范围。

- MATCH

	- 使用 MATCH 函数在 范围 单元格中搜索特定的项，然后返回该项在此区域中的相对位置。
	- MATCH(lookup_value, lookup_array, [match_type])

		- lookup_value    必需。 要在 lookup_array 中匹配的值。 例如，如果要在电话簿中查找某人的电话号码，则应该将姓名作为查找值，但实际上需要的是电话号码。

			- lookup_value 参数可以为值（数字、文本或逻辑值）或对数字、文本或逻辑值的单元格引用。

		- lookup_array    必需。 要搜索的单元格区域。
		- match_type    可选。 数字 -1、0 或 1。 match_type 参数指定 Excel 如何将 lookup_value 与 lookup_array 中的值匹配。 此参数的默认值为 1。

			- 1 或省略

				- MATCH 查找小于或等于 lookup_value 的最大值。 lookup_array 参数中的值必须以升序排序，例如：...-2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE。

			- 0

				- MATCH 查找完全等于 lookup_value 的第一个值。 lookup_array 参数中的值可按任何顺序排列。

			- -1

				- MATCH 查找大于或等于 lookup_value 的最小值。 lookup_array 参数中的值必须按降序排列，例如：TRUE, FALSE, Z-A, ...2, 1, 0, -1, -2, ... 等等。

- INDEX+MATCH

	- 反向查找

		- =INDEX(A2:A24,MATCH(E3,B2:B24))

	- 双向查找

		- =INDEX(B33:D36,MATCH(B40,A33:A36,0),MATCH(A40,B32:D32,0))

	- 多条件查找

		- =INDEX(C46:C49,MATCH(A47&B47,A46:A49&B46:B49,0))
