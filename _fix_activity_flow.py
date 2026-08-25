# -*- coding: utf-8 -*-
"""
_fix_activity_flow.py
扩展教学设计文档中教学组织表的R5/R6/R7内容（任务1/2/3）。
对29个偶数表(Table 6,8,10,...,62)的R5(知识讲解)、R6(技术演示)、R7(实操训练)
三个行的C2单元格写入800-1500字的详细教学内容。

每格内容包含：提问、概念分析、技术练习(含代码)、实验演示、结论、德育渗透、板书设计。
"""
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

# ============================================================
# 辅助函数
# ============================================================

def set_cell_multiline(tc, text, font_name='仿宋', font_size='21'):
    """将多行文本写入表格单元格，保留换行。"""
    paras = tc.findall(qn('w:p'))
    if not paras:
        p = OxmlElement('w:p')
        tc.append(p)
        paras = [p]
    p = paras[0]
    for r in p.findall(qn('w:r')):
        p.remove(r)
    lines = text.split('\n')
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    rPr.append(rFonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), font_size)
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), font_size)
    rPr.append(szCs)
    r.append(rPr)
    for li, line in enumerate(lines):
        if li > 0:
            br = OxmlElement('w:br')
            r.append(br)
        t_elem = OxmlElement('w:t')
        t_elem.set(qn('xml:space'), 'preserve')
        t_elem.text = line
        r.append(t_elem)
    p.append(r)
    for extra_p in paras[1:]:
        tc.remove(extra_p)


def get_tc_text(tc):
    """获取单元格纯文本。"""
    txt = ''
    for p in tc.findall(qn('w:p')):
        for r in p.findall(qn('w:r')):
            for elem in r:
                if elem.tag == qn('w:t'):
                    txt += elem.text or ''
    return txt.strip()


# ============================================================
# 任务配置数据
# ============================================================

# 章节映射: table_index -> (chapter_num, chapter_name)
CHAPTER_MAP = {
    6: (1, '初识数据分析'), 8: (1, '初识数据分析'),
    10: (1, '初识数据分析'), 12: (1, '初识数据分析'),
    14: (2, 'Excel数据分析工具'), 16: (2, 'Excel数据分析工具'),
    18: (2, 'Excel数据分析工具'), 20: (2, 'Excel数据分析工具'),
    22: (3, 'Numpy数学运算库'), 24: (3, 'Numpy数学运算库'),
    26: (3, 'Numpy数学运算库'),
    28: (4, 'Pandas数据分析库'), 30: (4, 'Pandas数据分析库'),
    32: (4, 'Pandas数据分析库'), 34: (4, 'Pandas数据分析库'),
    36: (5, 'SciPy科学计算库'), 38: (5, 'SciPy科学计算库'),
    40: (5, 'SciPy科学计算库'),
    42: (6, 'Sklearn数据统计基础'), 44: (6, 'Sklearn数据统计基础'),
    46: (6, 'Sklearn数据统计基础'),
    48: (7, 'Sklearn数据统计进阶'), 50: (7, 'Sklearn数据统计进阶'),
    52: (7, 'Sklearn数据统计进阶'),
    54: (8, 'Seaborn可视化分析库'), 56: (8, 'Seaborn可视化分析库'),
    58: (8, 'Seaborn可视化分析库'),
    60: (9, '综合评价与课程总结'), 62: (9, '综合评价与课程总结'),
}

# 每个table的R5/R6/R7子场景名 (从文档实际内容提取)
TASK_SUBS = {
    6:   ['认识数据分析', '常用数据分析方法', '综合练习'],
    8:   ['数据分析指标', '常用数据分析工具', '综合练习'],
    10:  ['数据分析方法理论', '任务实施', '综合练习'],
    12:  ['任务实施方法论', '综合练习', '成果制作与优化：初识数据分析'],
    14:  ['Excel概述', 'Excel数据分析技巧', '综合练习'],
    16:  ['函数相关概念', 'Excel图表可视化', '综合练习'],
    18:  ['Excel函数', '任务实施', '综合练习'],
    20:  ['任务实施方法论', '综合练习', '成果制作与优化：Excel数据分析工具'],
    22:  ['位运算函数', '统计函数', '综合练习'],
    24:  ['取反运算', '任务实施', '综合练习'],
    26:  ['任务实施方法论', '综合练习', '成果制作与优化：Numpy数学运算库'],
    28:  ['Pandas统计函数', '小数定标标准化数据', 'transform()函数与滚动窗口'],
    30:  ['标准化数据', 'apply()函数', '透视表'],
    32:  ['标准差标准化数据', 'agg()函数', '交叉表'],
    34:  ['任务实施方法论', '综合练习', '成果制作与优化：Pandas数据分析库'],
    36:  ['SciPy简介与安装', 'SciPy常用模块', '任务实施'],
    38:  ['cluster模块', '统计函数', '综合练习'],
    40:  ['任务实施方法论', '综合练习', '成果制作与优化：SciPy科学计算库'],
    42:  ['sklearn简介及安装', '数据集', '特征提取'],
    44:  ['sklearn安装与数据处理', '数据处理', '任务实施'],
    46:  ['任务实施方法论', '综合练习', '成果制作与优化：Sklearn数据统计基础'],
    48:  ['分类模型', '聚类模型', 'DBSCAN聚类算法'],
    50:  ['回归模型', 'K-MEANS聚类算法', '模型评估'],
    52:  ['任务实施方法论', '综合练习', '成果制作与优化：Sklearn数据统计进阶'],
    54:  ['seaborn概述与分类图', '回归图', '综合练习'],
    56:  ['关系图与分布图', '分布图与样式管理', '样式管理'],
    58:  ['任务实施方法论', '综合练习', '成果制作与优化：Seaborn可视化分析库'],
    60:  ['课程成果汇报与评价', '课程总结与复习', '综合练习'],
    62:  ['任务实施方法论', '综合练习', '成果制作与优化：综合评价与课程总结'],
}

# 行类型映射
ROW_TYPES = {5: '知识讲解', 6: '技术演示', 7: '实操训练'}


# ============================================================
# 代码模板（按章节）
# ============================================================

def get_code_example(chapter, sub_name, row_type):
    """根据章节和子场景返回代码示例。"""
    codes = {
        1: {
            '认识数据分析': (
                "  import pandas as pd\n"
                "  df = pd.read_csv('sales_data.csv')\n"
                "  print(df.head())          # 查看前5行\n"
                "  print(df.describe())      # 描述性统计\n"
                "  print(df.info())          # 数据类型与缺失值"
            ),
            '常用数据分析方法': (
                "  # 对比分析：本月vs上月\n"
                "  growth = (df[df['month']==11]['amount'].sum() - df[df['month']==10]['amount'].sum()) / df[df['month']==10]['amount'].sum() * 100\n"
                "  # 交叉分析：地区×品类\n"
                "  pivot = df.pivot_table(values='qty', index='region', columns='category', aggfunc='sum')"
            ),
            '数据分析指标': (
                "  UV = df['user_id'].nunique()  # 独立访客\n"
                "  GMV = df[df['action']=='order']['amount'].sum()\n"
                "  conv_rate = orders / UV * 100  # 转化率\n"
                "  avg_price = GMV / orders      # 客单价"
            ),
            '常用数据分析工具': (
                "  import pandas as pd, matplotlib.pyplot as plt\n"
                "  df = pd.read_excel('sales.xlsx')\n"
                "  monthly = df.groupby('month')['amount'].agg(['sum','mean'])\n"
                "  monthly['sum'].plot(kind='bar', title='月度销售额')"
            ),
            '数据分析方法理论': (
                "  # RFM模型\n"
                "  rfm = df.groupby('user_id').agg(\n"
                "    R=('date', lambda x: (df['date'].max()-x.max()).days),\n"
                "    F=('order_id','count'), M=('amount','sum'))\n"
                "  rfm['R_label'] = pd.qcut(rfm['R'], 4, labels=[4,3,2,1])"
            ),
        },
        2: {
            'Excel概述': (
                "  Excel公式：\n"
                "  =SUMIF(B:B,\"食品\",D:D)      '按品类求和\n"
                "  =VLOOKUP(A2,Sheet2!A:D,4,0) '查找匹配\n"
                "  =COUNTIF(C:C,\">100\")        '条件计数"
            ),
            'Excel数据分析技巧': (
                "  1.数据透视表：行=地区,列=品类,值=销售额(求和)\n"
                "  2.数据清洗：=TRIM(A2) =SUBSTITUTE(A2,\"-\",\"\")\n"
                "  3.条件格式：色阶(红-黄-绿)→热力图\n"
                "  4.分析工具库：数据→数据分析→描述统计"
            ),
            '函数相关概念': (
                "  =SUMIFS(D:D, B:B,\"食品\", C:C,\">100\")  '多条件求和\n"
                "  =INDEX(D:D, MATCH(A2, B:B, 0))           '反向查找\n"
                "  =IFS(D2>5000,\"VIP\", D2>2000,\"高级\", TRUE,\"普通\")  '多层判断\n"
                "  =SUMPRODUCT((B:B=\"食品\")*(D:D>100)*D:D)  '数组统计"
            ),
            'Excel图表可视化': (
                "  1.柱状图：选中品类和销售额→插入→柱状图\n"
                "  2.双轴图：主轴柱状(销售额)+次轴折线(增长率)\n"
                "  3.饼图：品类占比→添加百分比数据标签\n"
                "  4.散点图：价格×销量→添加趋势线和R²"
            ),
            'Excel函数': (
                "  =SUMIFS(D:D, B:B,\"食品\", C:C,\">100\")   '多条件求和\n"
                "  =INDEX(A:A, MATCH(\"苹果\", B:B, 0))        '反向查找\n"
                "  =IFS(D2>5000,\"VIP\", D2>2000,\"高级\", TRUE,\"普通\")\n"
                "  =SUMPRODUCT((B:B=\"食品\")*(C:C>100)*(D:D))  '多条件统计"
            ),
        },
        3: {
            '位运算函数': (
                "  import numpy as np\n"
                "  a = np.array([12, 25, 8], dtype=np.uint8)\n"
                "  b = np.array([10, 15, 6], dtype=np.uint8)\n"
                "  print(np.bitwise_and(a, b))  # [8, 9, 0]\n"
                "  print(np.bitwise_or(a, b))   # [14, 31, 14]\n"
                "  print(np.bitwise_xor(a, b))  # [6, 22, 14]\n"
                "  print(np.left_shift(a, 2))   # [48, 100, 32]"
            ),
            '统计函数': (
                "  sales = np.array([[100,150,120,200],[80,90,110,130],[200,180,220,250]])\n"
                "  print(np.sum(sales))              # 总额\n"
                "  print(np.mean(sales, axis=0))     # 月均\n"
                "  print(np.std(sales, axis=1))     # 波动\n"
                "  print(np.argmax(sales, axis=1))   # 最佳月\n"
                "  print(np.corrcoef(sales[0], sales[1]))  # 相关系数"
            ),
            '取反运算': (
                "  a = np.array([0, 5, 10, 255], dtype=np.uint8)\n"
                "  print(np.bitwise_not(a))    # [255, 250, 245, 0]\n"
                "  mask = np.array([True, False, True, False])\n"
                "  print(np.logical_not(mask)) # [False, True, False, True]\n"
                "  # 条件取反筛选\n"
                "  low = sales[~(sales > 100)]  # 取反获取低值"
            ),
        },
        4: {
            'Pandas统计函数': (
                "  print(df.describe())                          # 描述统计\n"
                "  print(df.groupby('category')['amount'].agg(['mean','count','std']))\n"
                "  print(df['category'].value_counts(normalize=True))  # 频次占比\n"
                "  print(df[['amount','qty','price']].corr())    # 相关性"
            ),
            '小数定标标准化数据': (
                "  def decimal_scaling(s):\n"
                "      j = np.ceil(np.log10(s.abs().max()))\n"
                "      return s / (10 ** j)\n"
                "  df_std = df.apply(decimal_scaling)\n"
                "  # 对比Min-Max: (df-df.min())/(df.max()-df.min())\n"
                "  # 对比Z-score: (df-df.mean())/df.std()"
            ),
            'transform()函数与滚动窗口': (
                "  df['group_mean'] = df.groupby('product')['sales'].transform('mean')\n"
                "  daily_ma7 = daily.rolling(window=7, min_periods=1).mean()\n"
                "  daily_std7 = daily.rolling(window=7).std()"
            ),
            '标准化数据': (
                "  from sklearn.preprocessing import MinMaxScaler, StandardScaler\n"
                "  mm = MinMaxScaler()\n"
                "  df_mm = pd.DataFrame(mm.fit_transform(df), columns=df.columns)\n"
                "  ss = StandardScaler()\n"
                "  df_zs = pd.DataFrame(ss.fit_transform(df), columns=df.columns)"
            ),
            'apply()函数': (
                "  df['level'] = df['price'].apply(lambda x: '高价' if x > 50 else '低价')\n"
                "  df['grade'] = df.apply(lambda r: '明星' if r['price']*r['qty']>5000 and r['rating']>=4.0 else '普通', axis=1)\n"
                "  df[['revenue','profit']] = df.apply(lambda r: pd.Series([r['price']*r['qty'], r['price']*r['qty']*0.3]), axis=1)"
            ),
            '透视表': (
                "  pt = pd.pivot_table(df, values='amount', index='region',\n"
                "                      columns='product', aggfunc='sum', fill_value=0, margins=True)\n"
                "  ct = pd.crosstab(df['region'], df['channel'], margins=True)"
            ),
            '标准差标准化数据': (
                "  df_zs = (df - df.mean()) / df.std()\n"
                "  from sklearn.preprocessing import StandardScaler\n"
                "  scaler = StandardScaler()\n"
                "  df_sk = pd.DataFrame(scaler.fit_transform(df), columns=df.columns)\n"
                "  # 验证: print(df_zs.mean()) → ≈0, print(df_zs.std()) → ≈1"
            ),
            'agg()函数': (
                "  # 多函数聚合\n"
                "  result = df.groupby('category')['amount'].agg(['sum','mean','count','std'])\n"
                "  # 不同列不同聚合\n"
                "  result2 = df.groupby('category').agg({'amount':'sum', 'qty':'mean', 'order_id':'count'})\n"
                "  # 命名聚合\n"
                "  result3 = df.groupby('category').agg(总额=('amount','sum'), 均价=('amount','mean'))"
            ),
            '交叉表': (
                "  ct = pd.crosstab(df['region'], df['category'], margins=True)\n"
                "  ct_pct = pd.crosstab(df['region'], df['category'], normalize='index')  # 行百分比\n"
                "  ct_all = pd.crosstab([df['region'],df['channel']], df['category'])  # 多级"
            ),
        },
        5: {
            'SciPy简介与安装': (
                "  # 安装SciPy\n"
                "  pip install scipy\n"
                "  # 导入子模块\n"
                "  from scipy import stats, cluster, optimize, integrate\n"
                "  import numpy as np\n"
                "  # 基础统计\n"
                "  data = np.random.normal(100, 15, 1000)\n"
                "  print(stats.describe(data))  # 描述性统计"
            ),
            'SciPy常用模块': (
                "  from scipy import stats, optimize, interpolate\n"
                "  # 正态性检验\n"
                "  print(stats.shapiro(data))  # W值, p值\n"
                "  # 优化求解\n"
                "  result = optimize.minimize(lambda x: x**2+3*x+2, x0=0)\n"
                "  # 插值\n"
                "  f = interpolate.interp1d(x, y, kind='cubic')"
            ),
            'cluster模块': (
                "  from scipy.cluster.hierarchy import linkage, fcluster, dendrogram\n"
                "  import matplotlib.pyplot as plt\n"
                "  Z = linkage(data, method='ward')\n"
                "  labels = fcluster(Z, t=3, criterion='maxclust')\n"
                "  dendrogram(Z)\n"
                "  plt.savefig('dendrogram.png')"
            ),
            '统计函数': (
                "  from scipy import stats\n"
                "  # t检验\n"
                "  t_stat, p_val = stats.ttest_ind(group1, group2)\n"
                "  # 卡方检验\n"
                "  chi2, p, dof, exp = stats.chi2_contingency(ct)\n"
                "  # 正态分布概率\n"
                "  prob = stats.norm.cdf(1.96)  # ≈0.975"
            ),
        },
        6: {
            'sklearn简介及安装': (
                "  pip install scikit-learn\n"
                "  from sklearn import datasets, preprocessing, model_selection\n"
                "  # 加载内置数据集\n"
                "  iris = datasets.load_iris()\n"
                "  X, y = iris.data, iris.target\n"
                "  print(X.shape, y.shape)\n"
                "  print(iris.feature_names)"
            ),
            '数据集': (
                "  from sklearn import datasets\n"
                "  from sklearn.model_selection import train_test_split\n"
                "  iris = datasets.load_iris()\n"
                "  X_train, X_test, y_train, y_test = train_test_split(\n"
                "      iris.data, iris.target, test_size=0.3, random_state=42)\n"
                "  print(f'训练集:{X_train.shape}, 测试集:{X_test.shape}')"
            ),
            '特征提取': (
                "  from sklearn.feature_extraction.text import TfidfVectorizer\n"
                "  from sklearn.feature_extraction.text import CountVectorizer\n"
                "  docs = ['电商数据分析', 'Python数据分析', '数据可视化']\n"
                "  vectorizer = TfidfVectorizer()\n"
                "  X = vectorizer.fit_transform(docs)\n"
                "  print(vectorizer.get_feature_names_out())\n"
                "  print(X.toarray())"
            ),
            'sklearn安装与数据处理': (
                "  from sklearn.preprocessing import StandardScaler, LabelEncoder\n"
                "  scaler = StandardScaler()\n"
                "  X_scaled = scaler.fit_transform(X)\n"
                "  le = LabelEncoder()\n"
                "  y_encoded = le.fit_transform(y)\n"
                "  print(f'标准化后均值:{X_scaled.mean():.4f}, 标准差:{X_scaled.std():.4f}')"
            ),
            '数据处理': (
                "  from sklearn.preprocessing import StandardScaler, MinMaxScaler\n"
                "  from sklearn.impute import SimpleImputer\n"
                "  # 缺失值填充\n"
                "  imputer = SimpleImputer(strategy='mean')\n"
                "  X_filled = imputer.fit_transform(X)\n"
                "  # 标准化\n"
                "  scaler = StandardScaler()\n"
                "  X_scaled = scaler.fit_transform(X_filled)"
            ),
        },
        7: {
            '分类模型': (
                "  from sklearn.linear_model import LogisticRegression\n"
                "  from sklearn.tree import DecisionTreeClassifier\n"
                "  from sklearn.metrics import accuracy_score, classification_report\n"
                "  clf = LogisticRegression(max_iter=200)\n"
                "  clf.fit(X_train, y_train)\n"
                "  y_pred = clf.predict(X_test)\n"
                "  print(f'准确率: {accuracy_score(y_test, y_pred):.2%}')\n"
                "  print(classification_report(y_test, y_pred))"
            ),
            '聚类模型': (
                "  from sklearn.cluster import KMeans, DBSCAN, AgglomerativeClustering\n"
                "  from sklearn.metrics import silhouette_score\n"
                "  km = KMeans(n_clusters=3, random_state=42)\n"
                "  labels = km.fit_predict(X)\n"
                "  print(f'轮廓系数: {silhouette_score(X, labels):.3f}')\n"
                "  print(f'聚类中心:\\n{km.cluster_centers_}')"
            ),
            'DBSCAN聚类算法': (
                "  from sklearn.cluster import DBSCAN\n"
                "  from sklearn.preprocessing import StandardScaler\n"
                "  X_scaled = StandardScaler().fit_transform(X)\n"
                "  db = DBSCAN(eps=0.5, min_samples=5)\n"
                "  labels = db.fit_predict(X_scaled)\n"
                "  n_clusters = len(set(labels)) - (1 if -1 in labels else 0)\n"
                "  print(f'聚类数: {n_clusters}, 噪声点: {sum(labels==-1)}')"
            ),
            '回归模型': (
                "  from sklearn.linear_model import LinearRegression, Ridge\n"
                "  from sklearn.metrics import r2_score, mean_squared_error\n"
                "  reg = LinearRegression()\n"
                "  reg.fit(X_train, y_train)\n"
                "  y_pred = reg.predict(X_test)\n"
                "  print(f'R²: {r2_score(y_test, y_pred):.3f}')\n"
                "  print(f'系数: {reg.coef_}, 截距: {reg.intercept_}')"
            ),
            'K-MEANS聚类算法': (
                "  from sklearn.cluster import KMeans\n"
                "  import matplotlib.pyplot as plt\n"
                "  inertias = []\n"
                "  for k in range(1, 11):\n"
                "      km = KMeans(n_clusters=k, random_state=42).fit(X)\n"
                "      inertias.append(km.inertia_)\n"
                "  plt.plot(range(1,11), inertias, marker='o')  # 肘部法则\n"
                "  plt.savefig('elbow.png')"
            ),
            '模型评估': (
                "  from sklearn.metrics import (accuracy_score, precision_score,\n"
                "      recall_score, f1_score, confusion_matrix, roc_auc_score)\n"
                "  print(f'准确率: {accuracy_score(y_test, y_pred):.3f}')\n"
                "  print(f'精确率: {precision_score(y_test, y_pred, average=\"macro\"):.3f}')\n"
                "  print(f'F1: {f1_score(y_test, y_pred, average=\"macro\"):.3f}')\n"
                "  print(confusion_matrix(y_test, y_pred))"
            ),
        },
        8: {
            'seaborn概述与分类图': (
                "  import seaborn as sns\n"
                "  import matplotlib.pyplot as plt\n"
                "  # 分类图\n"
                "  sns.barplot(data=df, x='category', y='amount')\n"
                "  plt.title('各品类销售额')\n"
                "  # 箱线图\n"
                "  sns.boxplot(data=df, x='category', y='amount')\n"
                "  # 小提琴图\n"
                "  sns.violinplot(data=df, x='category', y='amount')\n"
                "  plt.savefig('category_plots.png', dpi=150)"
            ),
            '回归图': (
                "  import seaborn as sns\n"
                "  # 回归图\n"
                "  sns.regplot(data=df, x='price', y='qty')\n"
                "  plt.title('价格-销量回归分析')\n"
                "  # 多变量回归\n"
                "  sns.lmplot(data=df, x='price', y='qty', hue='category', col='region')\n"
                "  plt.savefig('regression.png', dpi=150)"
            ),
            '关系图与分布图': (
                "  # 关系图\n"
                "  sns.relplot(data=df, x='price', y='qty', hue='category', size='amount')\n"
                "  # 分布图\n"
                "  sns.histplot(data=df, x='amount', kde=True, bins=30)\n"
                "  sns.kdeplot(data=df, x='amount', hue='category', fill=True)\n"
                "  plt.savefig('dist_plots.png', dpi=150)"
            ),
            '分布图与样式管理': (
                "  # 联合分布图\n"
                "  sns.jointplot(data=df, x='price', y='qty', kind='hex')\n"
                "  # 配对图\n"
                "  sns.pairplot(df[['price','qty','amount','rating']], hue='category')\n"
                "  # 样式管理\n"
                "  sns.set_style('whitegrid')\n"
                "  sns.set_context('talk')\n"
                "  sns.set_palette('Set2')"
            ),
            '样式管理': (
                "  # 五种主题样式\n"
                "  for style in ['darkgrid','whitegrid','dark','white','ticks']:\n"
                "      sns.set_style(style)\n"
                "      sns.barplot(data=df, x='category', y='amount')\n"
                "      plt.title(f'Style: {style}')\n"
                "      plt.savefig(f'style_{style}.png', dpi=100)\n"
                "      plt.clf()\n"
                "  # 自定义配色\n"
                "  sns.set_palette('husl', n_colors=8)"
            ),
        },
        9: {
            '课程成果汇报与评价': (
                "  # 综合项目评估脚本\n"
                "  import pandas as pd\n"
                "  projects = pd.DataFrame({\n"
                "    'group': ['A','B','C','D'],\n"
                "    'code_score': [28, 25, 26, 24],\n"
                "    'report_score': [18, 17, 19, 16],\n"
                "    'viz_score': [18, 15, 17, 16],\n"
                "    'defense_score': [9, 8, 9, 7]\n"
                "  })\n"
                "  projects['total'] = projects.iloc[:,1:].sum(axis=1)\n"
                "  print(projects.sort_values('total', ascending=False))"
            ),
            '课程总结与复习': (
                "  # 知识体系回顾\n"
                "  chapters = {\n"
                "    '第一章': '初识数据分析 - 流程/方法/指标',\n"
                "    '第二章': 'Excel - 透视表/函数/图表',\n"
                "    '第三章': 'Numpy - 位运算/统计/向量化',\n"
                "    '第四章': 'Pandas - 统计/标准化/透视/交叉',\n"
                "    '第五章': 'SciPy - cluster/stats/优化',\n"
                "    '第六章': 'Sklearn基础 - 数据集/特征/预处理',\n"
                "    '第七章': 'Sklearn进阶 - 分类/聚类/回归',\n"
                "    '第八章': 'Seaborn - 分类图/回归图/分布图/样式'\n"
                "  }"
            ),
        }
    }
    return codes.get(chapter, {}).get(sub_name, "  # 参见教材对应章节代码示例")


# ============================================================
# 德育渗透模板（按章节）
# ============================================================

def get_moral(chapter, sub_name, row_type):
    """根据章节返回德育渗透内容。"""
    morals = {
        1: "数据是国家战略资源，培养数据安全意识和数据伦理观念。分析数据时必须尊重用户隐私，遵守《个人信息保护法》。培养实事求是、严谨求实的科学态度，杜绝数据造假。",
        2: "工欲善其事必先利其器。掌握基础工具体现职业素养。数据表规范体现严谨的工作态度，函数学习需要耐心和逻辑思维，培养精益求精的工匠精神。",
        3: "位运算体现计算机科学的底层思维，从二进制理解数据本质，培养透过现象看本质的科学精神。向量化计算体现高效利用资源的理念，培养绿色计算和资源节约意识。",
        4: "标准化体现公平比较原则，如同社会公平需要消除不平等条件。移动平均体现辩证看待数据的思想，培养沉稳冷静的分析态度。多维交叉分析培养系统思维和全面分析能力。",
        5: "科学计算追求精确，培养严谨求实的科学精神。统计检验要客观公正，不选择性呈现结果。优化求解体现追求最优解的创新意识，培养持续改进的工作态度。",
        6: "机器学习服务社会，但需关注算法公平性和可解释性。数据预处理要保证质量，培养防患于未然的质量意识。特征工程体现从现象到本质的抽象能力，培养科学思维。",
        7: "模型评估要客观诚实，不可为追求高准确率而过度拟合。聚类分析体现物以类聚的哲学思想。回归分析培养因果推理能力，相关不等于因果，培养严谨的逻辑思维。",
        8: "数据可视化要求诚实呈现，不可通过调整比例误导读者，培养诚信意识和数据伦理。审美意识培养追求卓越的品质，细节决定品质，培养工匠精神。",
        9: "课程总结培养反思和自我评价能力。知识体系回顾培养系统思维和全局观念。成果汇报培养表达能力和技术自信，互评培养批判性思维和尊重他人成果的品格。",
    }
    return morals.get(chapter, "培养数据安全意识、工匠精神和团队协作能力，体现职业操守与社会责任。")


# ============================================================
# 成果制作任务专用内容
# ============================================================

def build_deliverable_content(table_idx, row, chapter_num, chapter_name, sub_name):
    """构建成果制作类任务的R5/R6/R7内容。"""
    tool = {1: 'Python Pandas', 2: 'Excel', 3: 'NumPy', 4: 'Pandas',
            5: 'SciPy', 6: 'Sklearn', 7: 'Sklearn', 8: 'Seaborn', 9: '综合工具'}.get(chapter_num, 'Python')
    if row == 5:
        return (
            f"任务1：任务实施方法论（知识讲解）\n"
            f"一、提问\n"
            f"  Q1: {chapter_name}成果应包含哪些内容？如何评价成果质量？\n"
            f"  Q2: 成果制作的标准流程是什么？各环节时间如何分配？\n"
            f"  Q3: 代码规范性、数据准确性和文档完整性如何保证？\n"
            f"二、概念分析\n"
            f"  {chapter_name}阶段成果制作是本章的综合实践环节。成果包括三部分："
            f"分析代码（Jupyter Notebook/Excel模板，可独立运行）、分析报告（Word文档，"
            f"包含背景/方法/结果/结论/建议五段式结构）、可视化图表（PNG格式，"
            f"分辨率150dpi以上，标题/坐标轴/图例完整）。评价标准采用五维评分："
            f"代码正确性（30%，无错误值/可运行/逻辑正确）、分析方法合理性（25%，"
            f"方法选择恰当/参数配置合理/多方法配合）、可视化质量（20%，图表类型匹配/"
            f"标题完整/配色美观）、结论与建议（15%，基于数据/有可操作性/有创新性）、"
            f"文档规范性（10%，结构清晰/注释完整/格式统一）。标准流程为："
            f"需求确认→数据准备→分析实施→可视化→报告撰写→检查优化→成果提交。\n"
            f"三、技术练习\n"
            f"  # 成果检查清单\n"
            f"  # 1.代码可运行性：逐个Cell执行，确认无报错\n"
            f"  # 2.数据准确性：交叉验证（透视表vs函数/手动抽样核对）\n"
            f"  # 3.图表完整性：标题/坐标轴/图例/数据标签\n"
            f"  # 4.报告结构：背景→方法→结果→结论→建议\n"
            f"  # 5.异常处理：IFERROR/try-except覆盖\n"
            f"  # 6.代码注释：关键步骤有中文注释\n"
            f"四、实验演示\n"
            f"  1. 讲解成果评价标准和五维评分体系\n"
            f"  2. 展示优秀成果范例，逐项分析其优点\n"
            f"  3. 演示代码检查方法（Ctrl+~显示公式/逐Cell执行）\n"
            f"  4. 演示数据交叉验证（透视表对比函数结果）\n"
            f"  5. 讲解报告撰写要点：逻辑清晰、图文并茂、数据支撑\n"
            f"  6. 学生开始规划自己的成果方案和时间分配\n"
            f"五、结论\n"
            f"  成果制作是学习成果的综合体现。好的分析成果应逻辑清晰、数据准确、"
            f"可视化直观、建议可行。标准化流程确保成果质量，五维评分提供客观评价依据。"
            f"成果要经得起检验，每个数字都可追溯，每个结论都有数据支撑。\n"
            f"六、德育渗透\n"
            f"  培养精益求精的工匠精神和质量意识。成果要经得起检验，"
            f"数据真实可靠，杜绝抄袭和弄虚作假，体现诚信品质和学术道德。"
            f"尊重他人知识产权，引用数据要标注来源。\n"
            f"七、板书设计\n"
            f"  评价标准：代码30%|方法25%|可视化20%|结论15%|规范10%\n"
            f"  制作流程：需求→数据→分析→可视化→报告→检查→提交\n"
            f"  成果清单：代码(.ipynb/.xlsx)+报告(.docx)+图表(.png)\n"
            f"  检查清单：可运行✓|准确✓|完整✓|规范✓|创新✓"
        )
    elif row == 6:
        return (
            f"任务2：综合练习（实操训练）\n"
            f"一、提问\n"
            f"  Q1: 如何在规定时间内完成完整的{chapter_name}分析成果？\n"
            f"  Q2: 成果检查时应关注哪些常见问题？如何避免？\n"
            f"  Q3: 如何提升成果的创新性和可操作性？\n"
            f"二、概念分析\n"
            f"  综合练习要求学生独立完成{chapter_name}成果的制作。要求在90分钟内完成："
            f"数据加载清洗（处理缺失值/异常值/重复值）、至少3种分析方法应用（描述统计+"
            f"对比分析+交叉分析或专项方法）、3张以上可视化图表（柱状图+折线图+饼图等）、"
            f"500字分析报告（含结论和建议）。常见问题包括：数据清洗不彻底导致统计偏差、"
            f"图表缺少标题和坐标轴标签影响可读性、结论过于简单缺乏数据支撑、"
            f"代码注释缺失降低可维护性、变量命名不规范影响可读性。"
            f"优秀成果特征：分析方法多元配合、可视化图表美观专业、"
            f"结论有深度洞察、建议有可操作性。\n"
            f"三、技术练习\n"
            f"  # 综合练习标准模板\n"
            f"  # Step1: 数据加载与清洗\n"
            f"  #   - pd.read_csv() / pd.read_excel()\n"
            f"  #   - dropna() / drop_duplicates() / fillna()\n"
            f"  # Step2: 多角度分析（至少3种方法）\n"
            f"  #   - describe() 描述性统计\n"
            f"  #   - groupby+agg 分组聚合\n"
            f"  #   - pivot_table 交叉分析\n"
            f"  # Step3: 可视化（至少3张图表）\n"
            f"  #   - 柱状图/折线图/饼图/散点图\n"
            f"  # Step4: 报告撰写（500字以上）\n"
            f"  #   - 背景/方法/结果/结论/建议\n"
            f"  # Step5: 检查优化\n"
            f"  #   - 交叉验证/图表美化/注释补充\n"
            f"四、实验演示\n"
            f"  1. 学生独立完成数据清洗和加载（20分钟）\n"
            f"  2. 实施3种分析方法，计算关键指标（25分钟）\n"
            f"  3. 制作3张可视化图表，确保标题完整（20分钟）\n"
            f"  4. 撰写500字分析报告，包含结论和建议（20分钟）\n"
            f"  5. 检查优化：交叉验证/代码注释/图表美化（5分钟）\n"
            f"  6. 教师巡回检查，记录常见问题并指导\n"
            f"五、结论\n"
            f"  综合练习检验了学生对{chapter_name}全流程的掌握程度。"
            f"时间管理是关键，建议按2:3:2:3分配清洗、分析、可视化、报告的时间。"
            f"常见问题可通过检查清单系统性避免。优秀成果需要在准确基础上追求创新和美观。\n"
            f"六、德育渗透\n"
            f"  培养时间管理能力和抗压能力。在有限时间内保质保量完成任务体现职业素养。"
            f"遇到困难不放弃，展现坚韧不拔的意志品质。成果要真实可靠，"
            f"不可为追求完美而篡改数据，培养诚信品质。\n"
            f"七、板书设计\n"
            f"  时间分配：清洗20'|分析25'|可视化20'|报告20'|检查5'\n"
            f"  常见问题：清洗不彻底/图表无标题/结论太简单/注释缺失\n"
            f"  检查清单：数据✓|分析✓|图表✓|报告✓|代码✓\n"
            f"  优秀特征：多元分析|专业可视化|深度洞察|可行建议"
        )
    else:  # row == 7
        return (
            f"成果制作与优化：{chapter_name}成果制作、检查与优化\n"
            f"一、成果整合\n"
            f"  1. 整理分析代码，确保每个文件/Notebook可独立运行\n"
            f"  2. 补充报告内容：背景→方法→结果→结论→建议五段式结构\n"
            f"  3. 确保可视化截图清晰（分辨率150dpi以上，标题/坐标轴/图例完整）\n"
            f"  4. 代码添加中文注释，关键步骤说明逻辑和参数含义\n"
            f"  5. 变量命名规范（英文有意义命名），代码缩进统一\n"
            f"二、成果展示\n"
            f"  · 每人/组展示3分钟，说明：分析目标→方法选择→代码逻辑→关键结果\n"
            f"  · 重点展示分析思路和发现的业务洞察，而非代码细节\n"
            f"  · 其他同学提问，展示者答辩，培养表达和应变能力\n"
            f"  · 展示评分：表达清晰度/逻辑性/答辩应对\n"
            f"三、互评与自评\n"
            f"  · 互评：按照五维评价标准给其他组打分，写出优点和改进建议\n"
            f"  · 自评：总结本组成果的亮点（3条）和不足（2条），制定改进计划\n"
            f"  · 教师点评：共性问题分析、优秀做法推广、改进方向指导\n"
            f"四、检查优化\n"
            f"  · 数据准确性：核对计算结果，用透视表交叉验证函数结果\n"
            f"  · 代码规范性：变量命名/注释/缩进/异常处理\n"
            f"  · 报告完整性：是否有结论和建议，是否基于数据事实\n"
            f"  · 可视化质量：标题/坐标轴/图例/配色/数据标签是否完整\n"
            f"  · 创新性：是否有独特分析角度或创新可视化方式\n"
            f"五、课后提交\n"
            f"  最终版提交学习通，截止本周日24:00\n"
            f"  提交清单：分析代码(.ipynb/.xlsx) + 分析报告(.docx) + 图表(.png)\n"
            f"  命名规范：学号_姓名_{chapter_name}成果\n"
            f"六、德育渗透\n"
            f"  成果展示培养表达能力和技术自信。互评培养批判性思维和尊重他人成果的品格。"
            f"检查优化体现精益求精的工匠精神，不满足于'能跑'，追求'优秀'。"
            f"学术诚信：所有数据真实，引用标注来源，杜绝抄袭代写。\n"
            f"七、板书设计\n"
            f"  评价表：代码30%|方法25%|可视化20%|结论15%|规范10%\n"
            f"  展示流程：目标→方法→代码→结果→答辩（3分钟）\n"
            f"  提交清单：代码+报告+图表（命名规范）\n"
            f"  优化方向：准确性→规范性→美观性→创新性"
        )


# ============================================================
# 内容生成主函数
# ============================================================

def generate_content(table_idx, row, sub_name, chapter_num, chapter_name):
    """生成单个单元格的详细教学内容。"""
    # 成果制作类任务特殊处理
    if '成果制作' in sub_name or (row == 5 and sub_name == '任务实施方法论'):
        return build_deliverable_content(table_idx, row, chapter_num, chapter_name, sub_name)

    row_type = ROW_TYPES[row]
    code = get_code_example(chapter_num, sub_name, row_type)
    moral = get_moral(chapter_num, sub_name, row_type)

    # 根据行类型生成不同侧重点的内容
    if row == 5:  # 知识讲解
        return (
            f"任务1：{sub_name}（知识讲解）\n"
            f"一、提问\n"
            f"  Q1: {sub_name}的核心概念是什么？在电商数据分析中如何应用？\n"
            f"  Q2: {sub_name}的关键参数、配置和注意事项有哪些？\n"
            f"  Q3: {sub_name}与其他分析方法或工具相比有何优势和局限？\n"
            f"二、概念分析\n"
            f"  {sub_name}是{chapter_name}章节的核心知识点。本节系统讲解其定义、基本原理、"
            f"分类体系和适用场景。{sub_name}在电商数据分析中具有广泛应用价值，"
            f"能够帮助学生理解数据处理和分析的基本方法论。重点掌握其工作原理和参数含义，"
            f"理解输入数据要求、处理逻辑和输出结果的解读方法。通过理论讲解配合电商案例演示，"
            f"帮助学生建立从概念到应用的系统性认知。注意区分{sub_name}与其他相关概念的异同，"
            f"明确其适用边界和局限性，为后续技术演示和实操训练奠定坚实的理论基础。\n"
            f"三、技术练习\n"
            f"{code}\n"
            f"四、实验演示\n"
            f"  1. 讲解{sub_name}的基本概念、定义和原理\n"
            f"  2. 演示代码示例，逐行解释关键步骤和参数含义\n"
            f"  3. 运行代码，观察并分析输出结果\n"
            f"  4. 修改关键参数，对比不同参数配置的效果差异\n"
            f"  5. 结合电商案例讨论实际应用场景和注意事项\n"
            f"  6. 引导学生提问讨论，深化对概念的理解\n"
            f"五、结论\n"
            f"  {sub_name}是{chapter_name}的关键技能点。理解概念原理是基础，掌握参数配置是核心，"
            f"灵活应用于实际电商分析问题是目标。理论学习为后续实践操作打下坚实基础，"
            f"需通过反复练习将理论内化为技能。注意概念之间的关联性，构建完整的知识体系。\n"
            f"六、德育渗透\n"
            f"  {moral}\n"
            f"  培养严谨求实的科学态度，概念学习要追根溯源，不可浅尝辄止。\n"
            f"七、板书设计\n"
            f"  {sub_name}核心概念图（定义→原理→分类→应用）\n"
            f"  关键参数：参数1|参数2|参数3 → 各自含义与取值范围\n"
            f"  应用场景：电商分析/用户运营/销售预测/数据清洗\n"
            f"  注意事项：适用条件|常见误区|与其他方法对比"
        )
    elif row == 6:  # 技术演示
        return (
            f"任务2：{sub_name}（技术演示）\n"
            f"一、提问\n"
            f"  Q1: 如何用代码实现{sub_name}的完整操作流程？\n"
            f"  Q2: 操作中常见的错误有哪些？如何排查和修复？\n"
            f"  Q3: 不同参数配置对分析结果有什么影响？如何选择最优参数？\n"
            f"二、概念分析\n"
            f"  本环节演示{sub_name}的完整技术操作流程。包括环境准备（安装库/导入模块）、"
            f"数据导入与预处理、核心操作执行、结果验证与输出、参数调优与对比。"
            f"通过逐步演示，让学生掌握从环境搭建到结果输出的全流程操作技能。"
            f"重点讲解每一步的关键操作要点、代码编写规范和常见问题处理方法。"
            f"演示中将穿插讲解代码调试技巧，如print调试、类型检查、逐步执行法等，"
            f"帮助学生建立独立排查问题的能力。\n"
            f"三、技术练习\n"
            f"{code}\n"
            f"四、实验演示\n"
            f"  1. 环境准备：启动开发环境，导入所需库和数据集\n"
            f"  2. 教师演示完整操作流程，学生跟随同步操作\n"
            f"  3. 逐步讲解每行代码的作用、参数含义和注意事项\n"
            f"  4. 演示常见错误（如类型错误/空值/越界）及排查修复方法\n"
            f"  5. 修改参数对比效果，讨论最优参数选择策略\n"
            f"  6. 巡回指导，解答学生疑问，纠正操作错误\n"
            f"五、结论\n"
            f"  {sub_name}的技术操作需要理解每个步骤的含义。环境准备是前提，"
            f"代码执行是核心，结果验证是保障，参数调优是进阶。常见错误的排查能力是实战关键，"
            f"建议养成先print检查数据再执行分析的良好习惯。代码要规范整洁，变量命名有意义。\n"
            f"六、德育渗透\n"
            f"  {moral}\n"
            f"  规范操作意识：每一步操作都要严谨认真，培养职业操作规范和安全意识。\n"
            f"七、板书设计\n"
            f"  操作流程：环境→导入→预处理→执行→验证→调优→输出\n"
            f"  常见错误：类型错误|空值异常|索引越界|参数无效\n"
            f"  排查方法：print调试→type()检查→逐步执行→查阅文档\n"
            f"  代码规范：命名规范|注释完整|缩进统一"
        )
    else:  # row == 7, 实操训练
        return (
            f"任务3：综合练习（实操训练）\n"
            f"一、提问\n"
            f"  Q1: 如何综合运用{sub_name}和本章所学知识解决实际电商分析问题？\n"
            f"  Q2: 独立完成分析任务时如何规划步骤和管理时间？\n"
            f"  Q3: 分析结果如何验证准确性并转化为业务建议？\n"
            f"二、概念分析\n"
            f"  综合练习要求学生独立运用{chapter_name}所学知识，完成一个完整的电商数据分析任务。"
            f"任务涵盖：数据加载与清洗（处理缺失值/异常值/重复值）、{sub_name}方法应用（核心分析）、"
            f"结果分析与可视化（图表制作）、结论输出与业务建议（报告撰写）。"
            f"核心是培养学生的独立分析和解决问题的能力，将课堂所学转化为实战技能。"
            f"要求学生在规定时间内独立完成，培养时间管理和抗压能力。"
            f"通过小组互助和展示分享，促进同伴学习和交流表达能力。\n"
            f"三、技术练习\n"
            f"{code}\n"
            f"  # 综合应用：独立完成完整分析流程\n"
            f"  # Step1: 加载数据并进行清洗（dropna/drop_duplicates）\n"
            f"  # Step2: 应用{sub_name}方法进行核心分析\n"
            f"  # Step3: 制作可视化图表展示分析结果\n"
            f"  # Step4: 检查结果正确性，优化代码和图表\n"
            f"  # Step5: 撰写分析结论和业务建议\n"
            f"四、实验演示\n"
            f"  1. 学生独立领取电商数据集和分析任务书\n"
            f"  2. 分析任务需求，制定解决步骤和时间分配\n"
            f"  3. 独立编写代码并调试运行，记录遇到的问题\n"
            f"  4. 检查结果正确性（交叉验证），优化代码和图表\n"
            f"  5. 撰写分析结论，提出至少2条业务改进建议\n"
            f"  6. 小组互助，展示分享（每组3分钟），教师点评\n"
            f"五、结论\n"
            f"  综合练习是检验学习效果的核心环节。独立完成分析任务需要将知识融会贯通，"
            f"从需求分析到结果输出形成完整的分析闭环。实践中遇到问题是正常现象，"
            f"关键是培养排查问题和解决问题的能力。分析结论要基于数据事实，建议要具有可操作性。"
            f"时间管理建议：数据清洗20%、分析编码30%、可视化20%、报告30%。\n"
            f"六、德育渗透\n"
            f"  {moral}\n"
            f"  培养独立解决问题和团队协作精神，遇到困难不放弃，展现坚韧不拔的意志品质。"
            f"分析结论要客观真实，杜绝数据造假。\n"
            f"七、板书设计\n"
            f"  实操流程：需求分析→数据清洗→{sub_name}→可视化→报告\n"
            f"  时间分配：清洗20%|编码30%|可视化20%|报告30%\n"
            f"  分析方法：{sub_name} + 本章关联知识综合运用\n"
            f"  评估标准：正确性/完整性/规范性/创新性"
        )


# ============================================================
# 主执行逻辑
# ============================================================

def main():
    fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
    doc = Document(fp)

    table_indices = list(range(6, 63, 2))  # 6, 8, 10, ..., 62 (共29个)
    updated = 0

    for idx, ti in enumerate(table_indices):
        seq = idx + 1
        t = doc.tables[ti]
        chapter_num, chapter_name = CHAPTER_MAP[ti]
        subs = TASK_SUBS[ti]

        for ri in [5, 6, 7]:
            sub_name = subs[ri - 5]
            row = t.rows[ri]
            tcs = row._tr.findall(qn('w:tc'))

            # 找到内容单元格（C2，即第3个tc，索引2）
            content_tc = None
            if len(tcs) > 2:
                content_tc = tcs[2]
            else:
                # 备选：找文本最长的tc
                best_tc = None
                best_len = 0
                for tc in tcs:
                    txt = get_tc_text(tc)
                    if len(txt) > best_len:
                        best_len = len(txt)
                        best_tc = tc
                content_tc = best_tc

            if content_tc is None:
                print(f'  [跳过] Table {ti} R{ri}: 未找到内容单元格')
                continue

            # 生成新内容
            new_content = generate_content(ti, ri, sub_name, chapter_num, chapter_name)
            old_text = get_tc_text(content_tc)
            old_len = len(old_text)

            # 写入新内容
            set_cell_multiline(content_tc, new_content)
            updated += 1
            print(f'  [更新] Table {ti} R{ri} (seq {seq}): {sub_name} | {old_len}字→{len(new_content)}字')

    # 保存
    doc.save(fp)
    print(f'\n完成！共更新 {updated} 个单元格。')
    print(f'文件已保存至: {fp}')


if __name__ == '__main__':
    main()
