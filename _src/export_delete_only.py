# -*- coding: utf-8 -*-
"""Export only explanatory copy from the Zhiwen teaching deck.

Literary excerpts, verse lines, source labels and quotation cards are deliberately
excluded.  The remaining text is rejoined by visual line so the transcript reads as
prose rather than as a dump of PowerPoint shapes.
"""
import re
import sys
from pathlib import Path

from pptx import Presentation


FOOTER = "《蛭文》教学读本 / 文本细读与写作"
SOURCE = re.compile(
    r"^《[^》]{1,30}[^/]{0,20}\s*/\s*[^/]{1,24}$"
    r"|^《[^》]{1,30}》(?:开篇|节选|跋|（节选）)?$"
)
DASH_SOURCE = re.compile(r"^——\s*[^—]{1,16}$")
EMU = 12700

PAGE_OVERRIDES = {
    2: (
        "先把“假”放下\n\n"
        "这不是一份需要背诵的教材，而是一部由学生写成的戏仿文集。它借课本的体例、"
        "经典的句法、庄重的文言，记录一所学校里被反复讲述的日常。\n\n"
        "阅读契约一：不作人身指认。文中的“徐规尔”“水蛭”是叙述者建构的讽刺形象，"
        "课堂讨论针对写法与经验，不针对现实个人。\n\n"
        "阅读契约二：不把戏谑当事实。夸张、传闻、诅咒都是文本策略，不等于证据；"
        "读它，要同时读它说了什么，以及它为何这样说。\n\n"
        "阅读契约三：不只看热闹。笑过之后，要追问它为什么笑、笑给谁看、笑完还剩下什么。\n\n"
        "整本书的褒贬都从语言出发：语言究竟用来连接人，还是用来设门槛。"
    ),
    3: (
        "一本课本的四个声部\n\n"
        "四个单元不是四种题材，而是四种“借壳”的方法：借古诗的腔、借小说的镜子、"
        "借古文的礼、借演讲的火。\n\n"
        "壹 · 古典诗歌的借壳：歌行、散曲、新月抒情、现代短诗，共 4 篇（pp. 1—10）。\n\n"
        "贰 · 现当代叙事：小说、书信、情诗、片段、寓言、戏剧，共 8 篇（pp. 11—44）。\n\n"
        "叁 · 文言戏仿：史传、小品、语录、侍坐、列传，共 6 篇（pp. 45—62）。\n\n"
        "肆 · 见证文体：杂文、散文诗、回忆、讲话、纪实，共 8 篇（pp. 63—109）。\n\n"
        "另有序 1 篇、跋 1 篇、古诗词诵读 3 首、制度时间线 1 页；合计 24 篇课文、3 首诗词与序跋。"
    ),
    5: (
        "序不是正文的附庸，而是第一道解释\n\n"
        "序位于正文与读者之间，像一道门槛：它不讲故事，却规定我们如何进入故事。"
        "热奈特把标题、序、跋、封面、题辞这类文字称为“副文本”，也就是正文的门厅。\n\n"
        "序的三种功能：\n\n"
        "定调：把仓促与错漏提前说出，降低读者的敌意。\n\n"
        "授权：借语言文字属于祖先遗产的说法，为全书争取文化合法性。\n\n"
        "邀约：把读者拉进同一个阵营，让阅读变成共同承担。\n\n"
        "编者开篇先承认仓促上马，这不是单纯的客气：它先把可能的错漏变成读者共同承担的风险，"
        "也为后面的尖锐提前争取谅解。自谦，是一种先发制人的修辞。\n\n"
        "理论关键词：副文本 paratext / 门槛 seuil / 框架 frame"
    ),
    11: (
        "越往后，笑越少，痛越具体\n\n"
        "四个单元的情绪不是平铺，而是逐级加压；到第四单元，讽刺变成证词。\n\n"
        "01 讽喻：金色，位置较轻，笑意与锋芒同时存在。\n\n"
        "02 冷峻：靛蓝，语气转冷，权力开始以规则的面貌出现。\n\n"
        "03 荒诞：灰绿，笑声仍在，但已经带出寒意。\n\n"
        "04 灼热：珊瑚，情绪最重，文字开始承担作证的责任。\n\n"
        "跋：克制，墨蓝；把愤怒收住，留下的不是轻松，而是判断。"
    ),
    12: (
        "目录\n\n"
        "从“怎么读”到“怎么写”。\n\n"
        "序与导读：02—11。\n\n"
        "壹 · 以古写今：13—25。\n\n"
        "贰 · 规则的冷光：27—44。\n\n"
        "叁 · 庄辞写荒诞：46—58。\n\n"
        "肆 · 字字泣血：60—81。\n\n"
        "古诗词诵读、跋与尾声：82—95。"
    ),
    13: (
        "壹\n\n以古写今\n\n第一单元 · 古典诗歌的戏仿\n\n"
        "学其体式而活其精神，以古写今。\n\n4 篇 · pp. 1—10"
    ),
    17: (
        "声音的四种暴力\n\n"
        "声音先于人物出现，制造压迫；它本身不是热闹，而是另一种不能退出。\n\n"
        "以乐写怒，把骂人写成演奏。\n\n"
        "大小、轻重、快慢，全是权力节奏。\n\n"
        "沉默不是安静，是被压住。\n\n"
        "声音的强度曲线"
    ),
    26: (
        "贰\n\n规则的冷光\n\n第二单元 · 现当代叙事\n\n"
        "权力最日常的样子，是把命令说成规则。\n\n8 篇 · pp. 11—44"
    ),
    75: (
        "把一个人的处分，变成一群人的事件\n\n"
        "一场针对个人的处分，在演讲里被推成群体事件。个体先站住，随后扩展为复数，"
        "最后被引向一个可见的方向。演讲的力量，正在于它让听者从旁观者变成参与者。\n\n"
        "个体：先让一个具体的人站住。\n\n"
        "复数：再把“我”扩成“我们”。\n\n"
        "未来：最后用“光明就在眼前”给出方向。\n\n"
        "最有力的一句话，往往也是说话者把自己一并押上去的那一句。"
    ),
    45: (
        "叁\n\n庄辞写荒诞\n\n第三单元 · 文言戏仿\n\n"
        "越像古文，越能照见今天的荒诞。\n\n6 篇 · pp. 45—62"
    ),
    83: (
        "短歌行·咏蛭：把壮心塞进一只蛭\n\n"
        "借体：借曹操的壮语，换成一个卑琐对象。\n\n"
        "反差：壮语越宏大，越显出升迁算计的可笑。\n\n"
        "收束：对脂与礼的贪求，被写成钻营的全部动力。\n\n"
        "典雅的外壳没有替人物遮丑，反而把丑放大了。"
    ),
    85: (
        "手在抖，枪膛是空的\n\n"
        "《蛭文》以这首歌作跋，不是要给出答案，而是让读者看清自己站在什么位置。\n\n"
        "阿塔曼不是胜利者，而是一个手在抖、枪里没有子弹、头顶悬着绞环的人。"
    ),
    86: (
        "黑夜将临，为什么仍然要开口\n\n"
        "歌里没有英雄式的保证。它先说黑夜会来，田野会消失，苍穹不再歌唱；然后才说白昼终来。"
        "希望被放在最冷的句子里。\n\n"
        "先写失去：不回避黑夜，也不假装它不存在。\n\n"
        "再写忍耐：忍耐不是顺从。\n\n"
        "最后开枪：枪响未必击碎什么，但必须发出声音。\n\n"
        "真正的勇气不是保证胜利，而是在不确定中仍然选择开口。"
    ),
    88: (
        "孩子们走进教室，像士兵开入兵营\n\n"
        "三个比喻把上学写成运输与编队：士兵、信件、列车。它们都不是坏东西，"
        "但它们都要求准时、整齐、可追踪。\n\n"
        "兵营：身体被编队，时间被统一。\n\n"
        "邮筒：每个人都是一个可投递的单位。\n\n"
        "轨道：路线预先铺好，偏离就会被发现。\n\n"
        "越平静的句式，越让人不安。"
    ),
    89: (
        "在绝望中反抗绝望\n\n"
        "跋没有把教育简单判为恶。它承认教育曾被相信、被期待，也承认那套秩序确实塑造了我们的日常；"
        "正因如此，反抗才不是一句轻飘飘的口号。\n\n"
        "信任：人们相信教育能改变命运。\n\n"
        "怀疑：但承诺并未兑现。\n\n"
        "反抗：不是相信明天一定更好，而是拒绝今天就这样。\n\n"
        "这不是同义反复，而是承认没有现成出路之后，仍然选择行动。"
    ),
    59: (
        "肆\n\n字字泣血\n\n第四单元 · 见证文体\n\n"
        "当戏仿不再好笑，文字就要出庭作证。\n\n8 篇 · pp. 63—109"
    ),
    95: (
        "把假课本合上\n\n把真的问题留下\n\n"
        "文学不替我们给出答案。它只训练我们辨认：谁在说话，谁被写成工具，谁还愿意作证。"
    ),
}

PAGE_OVERRIDES.update({
    6: (
        "序里的三个判断，决定了全书怎么写\n\n"
        "语言宣言：把语言从工具提升为生存条件。语言一旦只剩门槛功能，人就失去交流与创造的能力。\n\n"
        "荒诞宣言：现实比虚构更荒诞，所以戏仿不是夸张，而是取证。\n\n"
        "生命宣言：写作是自我校正，也是把人从指标里赎回来。\n\n"
        "三个宣言对应三种理论姿态：语言本体论、荒诞现实主义、主体伦理。\n\n"
        "序的末尾引法农，把人定义为宇宙和声中的一个“是”。这不是装饰性的题辞，而是全书的价值锚点。"
    ),
    8: (
        "从封面上取下来的六种情绪\n\n"
        "封面用珊瑚红与靛蓝对半切开，中间夹着一条近黑的竖缝。它不是单纯追求好看，"
        "而是先把人与制度分成两个区域。\n\n"
        "珊瑚砖红（#D56A58）：体温、愤怒、集体证词。\n\n"
        "靛蓝（#1B4F8B）：制度、监控、夜。\n\n"
        "灰绿（#649788）：古典、距离、反讽。\n\n"
        "陈金（#B09263）：记忆、尊严、旧纸。\n\n"
        "纸白（#F3EDE4）：留白、阅读、呼吸。\n\n"
        "墨蓝（#0C182B）：不可见的权力。\n\n"
        "用法：单元一陈金，单元二靛蓝，单元三灰绿，单元四珊瑚；黑色只做结构与句号。"
    ),
    18: (
        "同一条江，两种命运\n\n"
        "《蛭歌行》借《琵琶行》的江头场景，却把原本的同情改成了对峙：一边是被听见的漂泊者，"
        "一边是制造噪音的权力者。\n\n"
        "起点：原作从夜送客写起，戏仿把它移到夜学。\n\n"
        "声音：琵琶声有起伏，击鼓声只剩压迫。\n\n"
        "人物：原作写身世飘零，戏仿写权力在场。\n\n"
        "转折：原作由同病相怜连接彼此，戏仿则让座中人彼此怨怼。\n\n"
        "结尾：原作以泪收束，戏仿让喧闹继续留在场中。\n\n"
        "白居易把音乐写成命运，佚名把噪音写成权力。"
    ),
    19: (
        "短章如刀：一支散曲的全部力气\n\n"
        "曲牌：《朝天子》本是多用于抒情与讽刺的小令。\n\n"
        "俚语：用俗语化的绰号，把人物压成一个可笑的轮廓。\n\n"
        "侧写：不正面写老师多凶，只写学生的愁眉与惧怕。\n\n"
        "收束：以血流人逃的结局，把管理写成捕食。\n\n"
        "短，不是简单；是把所有铺垫压进一口气。"
    ),
    24: (
        "视角一翻，读者就失去了安全位置\n\n"
        "当“我”是水蛭，读者被迫站在加害者的视角里。这不是替谁辩护，而是让你看见权力如何自我叙述。\n\n"
        "自述：加害者开口，反而暴露其逻辑，把监视说成偶遇。\n\n"
        "对视：双方各说各话，两种语言互不相通。\n\n"
        "时间：抒情的时间被改写为捕食的时间。\n\n"
        "视角不是技巧，是伦理位置。最不舒服的阅读，往往最接近真相。"
    ),
    29: (
        "死亡不是终点，是叙述的入口\n\n"
        "死讯先到，具体的死亡却被推到记忆后面。叙述者用零散回忆把一个人重新拼起来，"
        "也把规则如何运作拼出来。\n\n"
        "延迟：记忆拒绝给出准确时刻，反而获得了叙述权。\n\n"
        "商品：师生关系被写成货架与交易。\n\n"
        "规则：规矩只约束别人，不约束制定者。\n\n"
        "所谓平凡的恶意，正因为它每天都在发生，才不容易被指认。"
    ),
    33: (
        "抒情如何被功利污染\n\n"
        "诗的开头仍然像一首情诗：松林、潮声、晚钟。可利益意象一出现，风景立刻变成了账本。\n\n"
        "起句：自然意象先把读者放进一个可信的抒情现场。\n\n"
        "换喻：金银进入凝视，爱情被换算成结算。\n\n"
        "声音：口号进入清晨的朗读，抒情被行政语占据。\n\n"
        "情诗的器官被换掉：眼睛看钱，喉咙发命令，爱只剩下占有的姿势。"
    ),
    35: (
        "权力不是一个人的脾气\n\n"
        "把坏人换成一个层级，问题就变了：每一层都在向下一层复制自己受过的压力。\n\n"
        "阶梯：压力沿职位向下流动。\n\n"
        "习惯：压力被内化成日常，久而久之，连反抗的念头也被磨平。\n\n"
        "操场：罚跑是权力最小的表演。\n\n"
        "所谓东区无考试，其实是一句反话：考试没有消失，只是变成了日常。"
    ),
    39: (
        "答案没有出现，它只是贴近了你\n\n"
        "拒绝先让世界变得透明，透明又让空无更明显。寻找的终点不是找到，"
        "而是承认自己一直站在对象的旁边。\n\n"
        "空：连雨和水蛭都消失以后，世界只剩干燥与透明。\n\n"
        "绕：教室、办公室、教务处形成闭环。\n\n"
        "近：最后左臂一疼，答案贴身出现。\n\n"
        "世界像一个圆：无论走向哪里，总会与注定相遇的事物重逢。"
    ),
    48: (
        "史传先写形貌，再让德行自己露底\n\n"
        "文章先写才能，再写舞台，最后落下判决。三项依次展开，一个人的权力路径就被写完了。\n\n"
        "出身：籍贯、性情、才能，按列传体例逐项交代。\n\n"
        "空间：不在埠口酒肆，只在庠序，权力有固定舞台。\n\n"
        "判决：最后四字把戒尺从教具变成刑具。\n\n"
        "古文越讲究雅，写出的现实越显得不雅；这种错位就是讽刺的发动机。"
    ),
    50: (
        "小品如何用夸张拆掉威严\n\n"
        "作者不直接说可笑，而是把动作放慢、放大，让读者自己看见一个管理者如何变成被围观的奇景。\n\n"
        "时令：春光与踏青先制造一个应该轻松的世界。\n\n"
        "动作：疾行、怒掷、躬伏，把权力写成失态。\n\n"
        "围观：众人仰首、议论，威严在目光里被拆开。\n\n"
        "夸饰不是失真，而是把平时不敢看的细节放大到无法回避。"
    ),
    52: (
        "格言的身体：命令如何变成口头禅\n\n"
        "禁令词反复出现，句子越来越短，语气越来越硬。人物不在场，声音却占满了空间。\n\n"
        "排比：连续不断的不许，把生活切成一条条禁令。\n\n"
        "反义：前面说休息为主，后面立刻接学习不可中断。\n\n"
        "回声：话被反复引用，最后不再需要说话的人。\n\n"
        "当一句话被重复到不需要解释，它就从命令变成了环境。"
    ),
    54: (
        "那个想去看春天的人，为什么被排除\n\n"
        "前三人的回答都能被学校折算成职位与成绩，只有曾腓谈的是歌咏、春风与闲适。"
        "于是最后那一声拒绝，成了整篇最冷的一句话。\n\n"
        "志：进士、秀才、教务，都是可被考核的将来。\n\n"
        "春：歌咏与闲适，无法被折成分数。\n\n"
        "怒：一声拒绝，把不能管理的生活赶出门。\n\n"
        "文章没有判谁对谁错，只让四种志向并排站着，读者自然看见谁被留下。"
    ),
    56: (
        "一间教室，如何被写成一座笼子\n\n"
        "《蛭》从众人的双休与“我”的独留起笔，先写时间，再写空间，最后写众人的表情。"
        "压抑不是一句判断，而是一整套生活细节。\n\n"
        "起句：一声叹词，把读者直接拉进集体处境。\n\n"
        "对照：别人的双休与我的独留，构成时间的不平等。\n\n"
        "收束：不能察不是无知，而是不需要看见。\n\n"
        "最后一句不是轻松，而是把愤怒收进了一个苦笑。"
    ),
    62: (
        "罪状不是结论，是语言病理\n\n"
        "空话、作态、无的放矢，看起来在骂人，实际在分析一种话语方式："
        "它为什么长、为什么吓人、为什么不看对象。\n\n"
        "第一条：话说得越长，信息反而越少。\n\n"
        "第二条：监控、处分、吼叫，都是为了让对方闭嘴。\n\n"
        "第三条：对家长、学生、卫士说同一种官话，就是不看对象。\n\n"
        "文章模仿八股，不是向八股投降，而是让八股在模仿中自己出丑。"
    ),
    64: (
        "排比不是重复，是把两种世界并排审判\n\n"
        "你们的岁月越写越具体，越像一张管理清单；我的岁月越写越开阔，越像一个人重新获得呼吸。\n\n"
        "你们的岁月：试卷、练习册、自习；班主任、科任教师、领导；总结反思、调研试题。\n\n"
        "我的岁月：山岳、圣殿、少年；思想、抱负、决心；岩石、溪流、广场上的球。\n\n"
        "当句式完全相同，读者才会发现：被对照的不是两种生活，而是两种价值尺度。"
    ),
    65: (
        "他们是谁？\n\n"
        "文章最后把镜头转向这套岁月培养出的下一代，不再只写自己，而是追问："
        "这套生活正在培养什么样的人。\n\n"
        "呓语：知识被念成没有对象的咒语。\n\n"
        "怀抱：价值在功利者的怀里长大。\n\n"
        "反问：结尾的追问把读者拉上审判席。\n\n"
        "散文诗的“我”不是私人抒情，而是一个拒绝被指标定义的人。"
    ),
    74: (
        "演讲的节奏：从质问到宣言\n\n"
        "开头连续追问，中间追问凭什么，最后把听众引向一个可见的方向。"
        "每一次重复都把听众往前推一步。\n\n"
        "呼告：直接对“水蛭”说话，制造现场对峙。\n\n"
        "反问：让听众无法保持旁观。\n\n"
        "短句：句子越短，越像口号，越适合齐声回应。\n\n"
        "演讲不是把道理说完整，而是把一群人推到同一个动作前面。"
    ),
    80: (
        "不是某个人坏，是结构在运转\n\n"
        "文章把“水蛭”写成结构中的一个位置：他也要向校长、向分数、向奖金负责。"
        "于是压迫不是脾气，而是可复制的流程。\n\n"
        "商品：学生的分数被折算成班主任的奖金。\n\n"
        "监控：监控像一双不会闭上的眼睛。\n\n"
        "规训：命令不再需要解释。\n\n"
        "把坏人改写成结构，不是替谁开脱，而是让读者看见：换一个人，机器仍然会转。"
    ),
    82: (
        "把杜甫的悲悯，借来写一次考试\n\n"
        "诗人借用《茅屋为秋风所破歌》的崇高体式，把个人的破屋之痛换成一次请假与休学。"
        "体式越悲壮，现实越显得荒诞。\n\n"
        "杜甫体：用最庄重的悲悯承载最日常的困境。\n\n"
        "以悲写小：宏大的声腔与小规模的现实互相拉扯。\n\n"
        "考试：整套戏仿最终落到学生的具体处境。\n\n"
        "读法：先辨认它借来的体裁，再找它把哪一处现实写坏了。"
    ),
})


def paragraph_text(shape):
    out = []
    for para in shape.text_frame.paragraphs:
        text = "".join(run.text for run in para.runs).strip()
        if text:
            out.append(text)
    return "\n".join(out)


def font_size(shape):
    sizes = []
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.font.size:
                sizes.append(run.font.size.pt)
    return max(sizes) if sizes else 0


def info(shape):
    return {
        "shape": shape,
        "text": paragraph_text(shape),
        "x": shape.left / EMU,
        "y": shape.top / EMU,
        "w": shape.width / EMU,
        "h": shape.height / EMU,
        "size": font_size(shape),
    }


def table_text(shape):
    rows = []
    for row in shape.table.rows:
        cells = [cell.text.strip().replace("\n", " / ") for cell in row.cells]
        if any(cells):
            rows.append(" | ".join(cells))
    return "\n".join(rows)


def filled_rects(slide):
    rects = []
    for shape in slide.shapes:
        if shape.has_table:
            continue
        if shape.has_text_frame and shape.text_frame.text.strip():
            continue
        if shape.shape_type != 1:
            continue
        try:
            fill = str(shape.fill.fore_color.rgb)
        except Exception:
            continue
        x, y = shape.left / EMU, shape.top / EMU
        w, h = shape.width / EMU, shape.height / EMU
        if 40 <= w <= 900 and 12 <= h <= 500 and w * h < 420000:
            rects.append((x, y, w, h, fill))
    return rects


def containing_rect(rects, item):
    cx = item["x"] + item["w"] / 2
    cy = item["y"] + item["h"] / 2
    inside = [r for r in rects if r[0] <= cx <= r[0] + r[2] and r[1] <= cy <= r[1] + r[3]]
    inside.sort(key=lambda r: r[2] * r[3])
    return inside[0] if inside else None


TABLE_PAGES = {12, 25, 44, 57, 81, 90, 93}
COLUMN_PAGES = {
    3, 6, 7, 8, 9, 14, 16, 20, 22, 24, 27, 29, 31, 33, 37, 41, 43,
    46, 48, 50, 52, 54, 56, 58, 60, 62, 64, 67, 68, 70, 72, 74, 77, 80,
    83, 84, 86, 88, 89, 91, 92,
}


def is_source(text, size=0, page=0):
    if len(text) > 48:
        return False
    if DASH_SOURCE.fullmatch(text):
        return True
    if not SOURCE.fullmatch(text):
        return False
    # Bare title references can be legitimate table cells (for example the
    # "底文" column).  Source labels are set much smaller than table body copy.
    if "/" in text:
        return True
    return page not in TABLE_PAGES and size <= 10.0


def is_quote_shape(page, item):
    """Page-specific geometry for direct quotations and verse blocks."""
    text, x, y, w, h, size = (item["text"], item["x"], item["y"], item["w"],
                              item["h"], item["size"])
    if page == 2:
        return x > 640 and 100 <= y <= 215
    if page == 4:
        return y < 230 and size >= 20
    if page == 5:
        return text.startswith("“本书从立项")
    if page == 6:
        return 160 <= y <= 205 and size < 13
    if page == 7:
        return size > 100
    if page == 17:
        return size >= 14 and x < 600 and y > 100
    if page in (21, 23):
        return 100 <= y < 450 and size <= 13.5
    if page == 35:
        return size >= 20 and x < 600 and y >= 150
    if page == 39:
        return size >= 20 and x < 600 and y >= 150
    if page == 51:
        return x > 600 and 130 <= y <= 380 and size >= 14
    if page in (52, 56, 74, 80, 84, 86):
        return x < 500 and y >= 190 and size >= 15
    if page == 55:
        return size >= 20 and x < 600 and y >= 150
    if page == 66:
        return x > 600 and 130 <= y <= 380 and size >= 14
    if page == 68:
        return x < 900 and (200 <= y <= 240 or 365 <= y <= 400)
    if page == 75:
        return size >= 20 and x < 600 and y >= 150
    if page == 78:
        return size >= 20 and x < 600 and y >= 150
    if page == 82:
        return x > 600 and 130 <= y <= 380 and size >= 14
    if page == 83:
        return x > 500 and w <= 60
    if page == 85:
        return y >= 120 and size >= 26 and x < 600
    if page == 88:
        return size >= 20 and x < 600 and y >= 150
    if page == 89:
        return size >= 20 and x < 600 and y >= 150
    if page == 94:
        return y >= 120 and size >= 26 and x < 600
    if page == 95:
        return False
    return False


def collect(slide, page):
    if page == 1:
        return (
            "蛭文\n\n一本假课本的读法\n\n"
            "从戏仿到见证：语言如何成为证词\n\n"
            "高中语文 · 文本细读与写作\n\n"
            "4 个单元 · 24 篇课文 · 3 首古诗词 · 序与跋"
        )

    items = []
    tables = []
    for shape in slide.shapes:
        if shape.has_table:
            text = table_text(shape)
            if text:
                tables.append({"text": text, "y": shape.top / EMU, "x": shape.left / EMU})
            continue
        if not shape.has_text_frame:
            continue
        item = info(shape)
        if not item["text"]:
            continue
        if item["y"] < 40 or item["text"] == FOOTER:
            continue
        if re.fullmatch(rf"{page:02d}", item["text"]) and item["y"] > 480:
            continue
        if re.fullmatch(r"\d{1,3} / 95", item["text"]):
            continue
        if re.fullmatch(r"\d{1,2}", item["text"]) and item["h"] > 50:
            continue
        items.append(item)

    source_items = [item for item in items if is_source(item["text"], item["size"], page)]
    rects = filled_rects(slide)
    remove = set()

    for src in source_items:
        rect = containing_rect(rects, src)
        if rect and len(source_items) <= 2:
            for item in items:
                if item is src:
                    continue
                cx = item["x"] + item["w"] / 2
                cy = item["y"] + item["h"] / 2
                if (rect[0] <= cx <= rect[0] + rect[2]
                        and rect[1] <= cy <= rect[1] + rect[3]
                        and item["y"] < src["y"] - 2
                        and item["size"] >= 9):
                    remove.add(id(item["shape"]))
    for item in items:
        if is_quote_shape(page, item):
            remove.add(id(item["shape"]))
        elif item["w"] <= 60 and item["h"] <= 30 and len(item["text"]) <= 1 and item["x"] > 500:
            remove.add(id(item["shape"]))

    remaining = [item for item in items if id(item["shape"]) not in remove]

    # Pull out the largest top-of-page heading; it is the page title in every layout.
    title = None
    candidates = [i for i in remaining if i["y"] < 105 and i["size"] >= 20]
    if candidates:
        title_item = max(candidates, key=lambda i: (i["size"], -i["y"], -i["x"]))
        title = title_item["text"]
        remove.add(id(title_item["shape"]))

    body = [item for item in remaining if id(item["shape"]) not in remove]

    # Join line boxes that are only visual continuations of one paragraph.  The
    # deck often interleaves the left and right columns by y-coordinate, so merge
    # inside x-columns first and order the finished blocks afterwards.
    columns = []
    for item in sorted(body, key=lambda i: i["x"]):
        if columns and abs(item["x"] - columns[-1][0]["x"]) <= 9:
            columns[-1].append(item)
        else:
            columns.append([item])
    merged = []
    for column in columns:
        column.sort(key=lambda i: i["y"])
        run = []
        for item in column:
            if run:
                prev = run[-1]
                same_column = abs(prev["x"] - item["x"]) <= 9 and abs(prev["w"] - item["w"]) <= 6
                close = 0 <= item["y"] - (prev["y"] + prev["h"]) <= 16
                min_size = 9.2 if page in TABLE_PAGES else 7.5
                same_type = abs(prev["size"] - item["size"]) <= 0.5 and item["size"] >= min_size
                open_tail = not prev["text"].rstrip().endswith(("。", "！", "？", "；", "："))
                if same_column and close and same_type and open_tail:
                    prev["text"] = prev["text"].rstrip() + item["text"].lstrip()
                    prev["h"] = max(prev["h"], item["y"] + item["h"] - prev["y"])
                    continue
            run.append(dict(item))
        merged.extend(run)
    leads = [i for i in merged if 90 <= i["y"] <= 180 and i["w"] >= 450]
    notes = [i for i in merged if i["y"] >= 350 and i["w"] >= 480]
    reserved = {id(i["shape"]) for i in leads + notes}
    core = [i for i in merged if id(i["shape"]) not in reserved]

    if page in COLUMN_PAGES:
        column_groups = []
        for item in sorted(core, key=lambda i: i["x"]):
            if column_groups and abs(item["x"] - column_groups[-1][0]["x"]) <= 40:
                column_groups[-1].append(item)
            else:
                column_groups.append([item])
        ordered = []
        for column in column_groups:
            ordered.extend(sorted(column, key=lambda i: i["y"]))
        core = ordered
    else:
        core.sort(key=lambda i: (i["y"], i["x"]))
    merged = sorted(leads, key=lambda i: (i["y"], i["x"])) + core + sorted(notes, key=lambda i: (i["y"], i["x"]))

    lines = []
    if title:
        lines.append(title)
    for item in merged:
        if item["text"] not in lines:
            lines.append(item["text"])
    for table in sorted(tables, key=lambda i: (i["y"], i["x"])):
        lines.append(table["text"])
    return "\n\n".join(lines)


def main():
    src = Path(sys.argv[1])
    dst = Path(sys.argv[2])
    pres = Presentation(str(src))
    chunks = []
    for i, slide in enumerate(pres.slides, 1):
        body = collect(slide, i)
        chunks.append(f"## 第{i:02d}页\n\n{body}")
    header = (
        "# 《蛭文》教学读本 · 讲解文字稿\n\n"
        f"共 {len(pres.slides)} 页。保留原有讲解、分析与出处，仅删去大段引文与诗作原文。\n"
    )
    dst.write_text(header + "\n\n---\n\n" + "\n\n---\n\n".join(chunks) + "\n", encoding="utf-8")
    print(f"wrote {dst} ({dst.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
