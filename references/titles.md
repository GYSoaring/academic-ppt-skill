# 标题：三种标题，三套规矩

新手最常犯的错，是把「论文标题」的写法套到「页标题」上。两者目标相反。

## 一、封面标题：可以有张力

来自 `nature-writing-skill` 的六定式（`references/02-title.md`），核心是**把约束写成贡献**：

| 定式 | 例子 | 何时用 |
|---|---|---|
| 工具名 + 冒号 + 机制 | `DynamicBind: predicting … with a deep generative model` | 方法论文最安全 |
| 陈述性发现 | `Highly accurate protein structure prediction with AlphaFold` | 结论稳、结果强 |
| 负定义 / 约束即贡献 | `Solving olympiad geometry without human demonstrations` | 你消除了一条传统必要条件 |
| 发现领头 | `Discovery of a structural class of antibiotics …` | 发现比方法重要 |
| 抱负式 | `Towards a general-purpose foundation model for …` | 未到通用，但想占坑 |

反模式：`A novel method for X` / `A novel deep learning framework for Y`——`novel` 不带信息量。

## 二、页标题：概括这一页讲什么，不是喊观点

**这是 deck 的默认模式（`title_mode: summary`）。**

- 通常 4–10 字，校验器硬性接受 4–18 字；名词短语，像论文的二级标题
- 读者只看标题，应该知道**这一页的对象与方法**，而不是被塞一个论断
- 判断放到正文第一句（`lead`）和 `claim` 字段里
- 不含判断词：是、不是、必须、决定、没有、无法、能、需要、不会

对照：

| 改成观点式（旧） | 改成概括式（新） |
|---|---|
| 逆设计本质是一个病态反问题 | 逆问题的形式与病态性 |
| 三条路线无法互相替代 | 三条路线的横向比较 |
| 仿真在环优化没有代理模型 | 仿真在环的参数优化 |
| 三个空白都没有独立判据 | 候选、置信度与制造约束 |

**为什么**：观点式标题连着排，读者看到的是**七句标语**，反而看不到论证；概括式标题给出地图，判断留在正文被论证支撑。

## 三、什么时候仍然用观点式

`title_mode: assertion` 是**可选**项，只在两种场合开：

- 答辩 / 路演：评委只看十秒，需要结论顶在脸上
- 汇报给决策者：他们只要结论，不要推导

开了就要付代价：标题会读起来像口号，正文必须立刻给出证据，否则整页就是空的。

## 四、面板小标题（三级）

面板标题是**名词或名词短语**，长度 2–6 字：

- 好：现象、机制、约束、读数、适用范围、验证口径
- 坏：看到什么、为什么、代价是什么、你需要知道

## 五、校验怎么跑

```bash
python scripts/validate_deck.py --deck deck.json --theme theme/theme.json --render-dir out/rendered --out out
```

- `title_mode: summary`：标题 4–18 字、不得含判断词、**每页必须有 `claim`**（判断不能既不在标题也不在正文）
- `title_mode: assertion`：标题必须含断言线索词

`qa_design.py` 在 summary 模式下自动跳过 `assertion_ratio`——不再拿「标题够不够冲」当分数。

## 六、自检

1. 把全部标题连起来读，是一张目录，还是一串口号？
2. 每页的 `claim` 是不是一句能被反对的话？不能反对 = 没说事。
3. 标题里的词，正文里都解释了吗？
