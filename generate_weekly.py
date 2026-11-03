# -*- coding: utf-8 -*-
"""
11111111111111
22222222222222
每周英语口语材料生成脚本
========================
每周使用前只需改【内容配置】区：N（周次）、THEME、LEVEL、TITLE、BODY、
GOAL、EXPRESSIONS、STRUCTURES、COMPREHENSION、SPEAKING、TRANSLATION。

然后运行：  python generate_weekly.py

会自动：
1. 生成 Word（沿用上一周 N-1 的排版）
2. 生成英音 / 美音两份 MP3（edge-tts，走本地代理）
3. 验证后原子重命名，清理临时文件

依赖：
- 系统 Python（有 python-docx）：用来生成 Word 和跑本脚本
- Anaconda Python（有 edge-tts）：用来生成音频
- 本地代理 127.0.0.1:7897（edge-tts 直连微软 TTS 会超时）

注意：如果目标周 N 的正式文件已存在，脚本会中止，绝不覆盖。
"""

import os
import subprocess
import hashlib

import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# ============================================================
# 内容配置（每周改这里）
# ============================================================
BASE_DIR = r'C:/Users/马江涛/OneDrive/桌面/work_daily/英语口语'

N = 17
THEME = 'Trying a New Restaurant'
LEVEL = 'A2+'
TITLE = 'Trying a New Restaurant'

BODY = [
    "Last Saturday, my friend Leo and I wanted to try a new restaurant near the park. We had heard good things about it, so we decided to go there for dinner. I usually cook at home on weekends, but this time I wanted to eat out and relax.",
    "We arrived at the restaurant around seven. It was busy, but a waiter found us a table near the window. He gave us the menu and asked if we wanted something to drink first. I said, \"I'd like a glass of water, please.\" Leo asked for some tea.",
    "The menu was long, so it took us a few minutes to choose. I wanted something light, so I decided to try the chicken salad. Leo said he needed to eat something warm because the evening was a little cold, so he ordered some soup and rice.",
    "Before we ordered, I asked the waiter if the chicken salad was spicy. He said it was not too spicy, but he could make it milder if I wanted. That made me feel sure about my choice. I also asked him what he would recommend, and he told us the fish was very popular.",
    "The food came after about fifteen minutes. Everything looked fresh and smelled wonderful. I think the salad was delicious because the vegetables were crisp and the chicken was tender. We talked about our week and ate slowly. It was nice to take our time and enjoy the meal.",
    "When we finished, I asked for the bill. The meal cost a little more than I expected, but it was worth it. I have learned that trying a new place can be fun, even if it is a bit more expensive. Next time, I want to bring my family there, because I think they would enjoy it too.",
]

GOAL = '能围绕去餐厅吃饭这一场景自然表达，学会点餐、询问、评价食物，并主动使用 want to、decide to、I\'d like、ask if、need to、I think...because 等高频口语结构组织句子。'

REMINDER = '特别提醒：别忘了自己阅读并录音。'

# 高频表达：[(英文, 中文, 例句), ...]
EXPRESSIONS = [
    ('eat out', '出去吃饭', 'I usually cook at home, but tonight I want to eat out.'),
    ('try a new place', '尝试一个新地方', 'We wanted to try a new place near the park.'),
    ('look at the menu', '看菜单', 'It took us a few minutes to look at the menu and choose.'),
    ('something to drink', '喝的东西', 'The waiter asked if we wanted something to drink first.'),
    ('a little cold', '有点冷', 'The evening was a little cold, so Leo ordered soup.'),
    ('take our time', '慢慢来，不着急', 'It was nice to take our time and enjoy the meal.'),
    ('ask for the bill', '买单，结账', 'When we finished, I asked for the bill.'),
    ('worth it', '值得', 'The meal cost a little more, but it was worth it.'),
    ('feel sure about', '对……有把握，放心', 'That made me feel sure about my choice.'),
    ('next time', '下次', 'Next time, I want to bring my family there.'),
]

# 口语结构：[(结构名, 中文作用, 原文例句英文, 原文例句中文, [(延伸英文, 延伸中文), ...]), ...]
STRUCTURES = [
    ('want to + 动词原形', '表示想要做某事。',
     'We wanted to try a new restaurant near the park.', '我们想试试公园附近的一家新餐厅。',
     [('I want to call my mom tonight.', '我今晚想给妈妈打电话。'),
      ('She wants to learn how to cook.', '她想学做饭。'),
      ('We want to take a walk after dinner.', '我们想晚饭后去散步。'),
      ('He wants to buy a new phone next month.', '他想下个月买一部新手机。')]),
    ('decide to + 动词原形', '表示决定做某事。',
     'I decided to try the chicken salad.', '我决定试试鸡肉沙拉。',
     [('I decided to go to bed early last night.', '我昨晚决定早点睡觉。'),
      ('She decided to take the bus to work.', '她决定坐公交去上班。'),
      ('We decided to cook dinner at home.', '我们决定在家做晚饭。'),
      ('They decided to meet after lunch.', '他们决定午饭后见面。')]),
    ('I\'d like + 名词', '表示我想要……，点餐或礼貌地提出请求时很常用。',
     'I\'d like a glass of water, please.', '请给我一杯水。',
     [('I\'d like a cup of coffee, please.', '请给我一杯咖啡。'),
      ('She\'d like a table for two.', '她想要一张两人的桌子。'),
      ('We\'d like some more time to choose.', '我们想多一点时间挑选。'),
      ('I\'d like the chicken salad, please.', '请给我来一份鸡肉沙拉。')]),
    ('ask someone if + 完整句子', '表示问某人是否……。',
     'I asked the waiter if the chicken salad was spicy.', '我问服务员鸡肉沙拉辣不辣。',
     [('I asked my friend if he was free on Saturday.', '我问朋友他周六是否有空。'),
      ('She asked me if I wanted some tea.', '她问我是否想喝点茶。'),
      ('We asked the driver if the bus stopped here.', '我们问司机这辆公交是否在这里停。'),
      ('He asked his manager if he could leave early.', '他问经理是否可以早点离开。')]),
    ('need to + 动词原形', '表示需要做某事。',
     'He needed to eat something warm because the evening was a little cold.', '他需要吃点热乎的，因为晚上有点冷。',
     [('I need to buy some food after work.', '我下班后需要买点吃的。'),
      ('She needs to finish her homework tonight.', '她今晚需要完成作业。'),
      ('We need to leave early tomorrow morning.', '我们明早需要早点出发。'),
      ('He needs to check the address again.', '他需要再核对一下地址。')]),
    ('I think + 观点 + because + 原因', '表达自己的观点，并说明原因。',
     'I think the salad was delicious because the vegetables were crisp and the chicken was tender.', '我觉得沙拉很好吃，因为蔬菜很脆，鸡肉很嫩。',
     [('I think this restaurant is good because the food is fresh.', '我觉得这家餐厅不错，因为食物很新鲜。'),
      ('I think cycling is relaxing because I can spend time outside.', '我觉得骑自行车很放松，因为我可以待在户外。'),
      ('I think English is useful because I can talk with more people.', '我觉得英语很有用，因为我能和更多人交流。'),
      ('I think this jacket is better because it fits me well.', '我觉得这件夹克更好，因为它很合身。')]),
    ('... more than I expected', '表示比我想象的更……。',
     'The meal cost a little more than I expected.', '这顿饭比我想象的贵了一点。',
     [('The movie was better than I expected.', '这部电影比我想象的好。'),
      ('The test was harder than I expected.', '这次考试比我想象的难。'),
      ('The bus came earlier than I expected.', '公交车比我想象的来得早。'),
      ('The city is bigger than I expected.', '这座城市比我想象的大。')]),
]

COMPREHENSION = [
    'Why did Leo and the writer decide to go to the new restaurant?',
    'What did the writer ask for to drink?',
    'Why did Leo order soup and rice?',
    'What did the waiter say about the chicken salad?',
    'How did the writer feel about the meal at the end?',
]

SPEAKING = [
    'Do you like trying new restaurants? Why or why not?',
    'What would you like to eat this weekend?',
    'When you go to a restaurant, what do you usually order?',
    'Have you ever tried a new place that you really liked? What did you order?',
    'Is there a dish you want to learn to cook? Why?',
]

# 训练安排：[(文本, 是否加粗), ...]
TRAINING = [
    ('周一、周二：精读文章并提取句型。弄懂文章意思、句子结构和常用表达，但不必背诵。重点学习“三、常用口语结构与延伸练习”，观察同一个句型怎样通过替换不同内容生成新的句子。每天选择其中 2—3 个结构，自己再口头造几个与真实生活有关的句子。', False),
    ('周三、周四：反复听。先不看原文听，再对照原文听。分别熟悉标准英音和标准美音。目标是听到一句话时能够较快理解它的意思，同时注意自己本周学习过的句型在真实语流中的发音。', False),
    ('周五、周六：跟读、模仿与替换造句。直接使用“一、英文原文”和英美两份音频进行练习。注意发音、停顿、重音、连读、弱读和语调。同时从本周口语结构中选择 3—5 个进行替换练习，例如更换人物、动作、地点、时间和原因，用同一个结构说出不同句子。', False),
    ('周日：主动输出。不看原文，用英语复述约 1—2 分钟，并尽可能自然使用本周学到的 3—5 个常用口语结构。说不完整也没关系，也可以结合自己的工作、生活、朋友、运动、周末计划等真实内容扩展表达。', False),
    ('特别注意：Word 中不需要、也不得另外生成“影子跟读版”或任何使用 /、// 标记停顿的正文副本。', True),
    ('本周训练的最终目标不是把文章背下来，而是把文章中的部分高频句型变成以后自己能够使用的表达工具。', True),
]

TRANSLATION = [
    '上周六，我和朋友 Leo 想去公园附近的一家新餐厅试试。我们听说了它的好评，于是决定去那里吃晚饭。我周末通常在家做饭，但这次我想出去吃，放松一下。',
    '我们大约七点到了餐厅。店里很忙，但一位服务员给我们找了一张靠窗的桌子。他把菜单递给我们，问我们想先喝点什么。我说：“请给我一杯水。”Leo 点了一些茶。',
    '菜单很长，所以我们花了几分钟才选好。我想吃点清淡的，于是决定试试鸡肉沙拉。Leo 说他需要吃点热乎的，因为晚上有点冷，于是他点了一份汤和米饭。',
    '点餐之前，我问服务员鸡肉沙拉辣不辣。他说不太辣，但如果我想，可以做得更清淡。这让我对自己的选择放心了。我还问他有什么推荐，他告诉我们鱼很受欢迎。',
    '大约十五分钟后，菜上来了。每样东西看起来都很新鲜，闻起来也很香。我觉得沙拉很好吃，因为蔬菜很脆，鸡肉很嫩。我们聊着这一周的事，慢慢吃。能慢慢享用这顿饭真好。',
    '吃完后，我要了账单。这顿饭比我想象的贵了一点，但很值得。我体会到，尝试一个新地方会很有趣，即使稍微贵一点。下次，我想带家人一起去，因为我觉得他们也会喜欢。',
]

# 音频配置
EDGE_PY = r'C:/anaconda/Anaconda/python.exe'
PROXY = 'http://127.0.0.1:7897'
VOICE_UK = 'en-GB-RyanNeural'
VOICE_US = 'en-US-ChristopherNeural'

# ============================================================
# 生成逻辑（一般不用改）
# ============================================================
OUT_DIR = os.path.join(BASE_DIR, str(N))
DOCX_PATH = os.path.join(OUT_DIR, f'{N}.docx')
MP3_UK = os.path.join(OUT_DIR, f'{N}_英音.mp3')
MP3_US = os.path.join(OUT_DIR, f'{N}_美音.mp3')
TMP_UK = os.path.join(OUT_DIR, f'{N}_英音.tmp.mp3')
TMP_US = os.path.join(OUT_DIR, f'{N}_美音.tmp.mp3')
AUDIO_TXT = os.path.join(OUT_DIR, '_audio_text.txt')


def fail(msg):
    print('[失败]', msg)
    raise SystemExit(1)


def check_no_overwrite():
    for p in (DOCX_PATH, MP3_UK, MP3_US):
        if os.path.exists(p):
            fail(f'正式文件已存在，拒绝覆盖：{p}')


def get_reference_docx():
    ref = os.path.join(BASE_DIR, str(N - 1), f'{N - 1}.docx')
    if os.path.exists(ref):
        return ref
    tpl = os.path.join(BASE_DIR, '模板.docx')
    if os.path.exists(tpl):
        return tpl
    fail(f'找不到排版参考文件（{ref} 或 模板.docx）')


def set_font(run, size_pt, bold):
    run.font.name = 'Times New Roman'
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn('w:eastAsia'), 'Times New Roman')
    run.font.size = Pt(size_pt)
    run.font.bold = bold


def build_word():
    doc = docx.Document(get_reference_docx())
    body_el = doc.element.body
    for child in list(body_el):
        if child.tag.split('}')[-1] == 'sectPr':
            continue
        body_el.remove(child)

    def add(text, size, bold=False, align=None, line=None, spB=None, spA=None):
        p = doc.add_paragraph()
        run = p.add_run(text)
        set_font(run, size, bold)
        pf = p.paragraph_format
        if align is not None:
            pf.alignment = align
        if line is not None:
            pf.line_spacing = line
        if spB is not None:
            pf.space_before = Pt(spB)
        if spA is not None:
            pf.space_after = Pt(spA)
        return p

    C = WD_ALIGN_PARAGRAPH.CENTER
    word_count = len(' '.join(BODY).split())

    add(f'第{N}周英语口语训练', 18, bold=True, align=C, spA=8)
    add(f'主题：{THEME}', 11, spA=2)
    add(f'难度：{LEVEL}', 11, spA=2)
    add(f'正文词数：{word_count}', 11, spA=2)
    add(f'本周目标：{GOAL}', 11, spA=2)
    add(REMINDER, 11, bold=True, spA=6)

    add('一、英文原文', 13, bold=True, spB=8, spA=5)
    add(TITLE, 14, bold=True, align=C, spA=7)
    for para in BODY:
        add(para, 12, line=1.5, spA=10)

    add('二、本周高频表达', 13, bold=True, spB=8, spA=5)
    for en, zh, ex in EXPRESSIONS:
        add(f'{en}：{zh}。例句：{ex}', 11, spA=3)

    add('三、常用口语结构与延伸练习', 13, bold=True, spB=8, spA=5)
    for name, role, ex_en, ex_zh, exts in STRUCTURES:
        add(f'结构：{name}', 11, bold=True, spA=1)
        add(f'中文作用：{role}', 10.5, spA=1)
        add(f'原文例句：{ex_en}', 10.5, spA=1)
        add(ex_zh, 10.5, spA=2)
        add('延伸例句：', 10.5, bold=True, spA=1)
        for j, (e_en, e_zh) in enumerate(exts, 1):
            add(f'{j}. {e_en}', 10.5, spA=2)
            add(e_zh, 10.5, spA=3)

    add('四、理解问题', 13, bold=True, spB=8, spA=5)
    for i, q in enumerate(COMPREHENSION, 1):
        add(f'{i}. {q}', 10.5, spA=2)

    add('五、口语输出问题', 13, bold=True, spB=8, spA=5)
    for i, q in enumerate(SPEAKING, 1):
        add(f'{i}. {q}', 10.5, spA=2)

    add('六、本周训练安排', 13, bold=True, spB=8, spA=5)
    for t, is_bold in TRAINING:
        spA = 3 if (is_bold and t.startswith('特别')) else 4
        add(t, 10.5, bold=is_bold, spA=spA)

    add('七、全文翻译', 13, bold=True, spB=8, spA=5)
    for t in TRANSLATION:
        add(t, 11, spA=7)

    doc.save(DOCX_PATH)
    print('[OK] Word 已生成：', DOCX_PATH)


def gen_audio(voice, out_tmp):
    audio_text = TITLE + '\n\n' + '\n\n'.join(BODY)
    with open(AUDIO_TXT, 'w', encoding='utf-8') as f:
        f.write(audio_text)
    cmd = [EDGE_PY, '-m', 'edge_tts', '--voice', voice, '--rate', '+0%',
           '--proxy', PROXY, '--file', AUDIO_TXT, '--write-media', out_tmp]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        fail(f'edge-tts 生成失败（{voice}）：\n' + (r.stderr or '')[-500:])


def mp3_duration(path):
    """粗略估算 MP3 时长（秒），解析失败返回 0。"""
    with open(path, 'rb') as f:
        data = f.read()
    size = len(data)
    br_mp3 = {1: 32, 2: 40, 3: 48, 4: 56, 5: 64, 6: 80, 7: 96, 8: 112,
              9: 128, 10: 160, 11: 192, 12: 224, 13: 256, 14: 320}
    br_mp3_2 = {0: 8, 1: 16, 2: 24, 3: 32, 4: 40, 5: 48, 6: 56, 7: 64,
                8: 80, 9: 96, 10: 112, 11: 128, 12: 144, 13: 160}
    i = 0
    while i < len(data) - 4:
        if data[i] == 0xFF and (data[i + 1] & 0xE0) == 0xE0:
            ver = (data[i + 1] >> 3) & 3
            layer = (data[i + 1] >> 1) & 3
            br_idx = (data[i + 2] >> 4) & 0xF
            if layer == 1:
                br = (br_mp3.get(br_idx) if ver == 3 else br_mp3_2.get(br_idx))
                if br:
                    return size * 8 / (br * 1000)
        i += 1
    return 0


def sha256(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def verify_audio(path):
    if not os.path.exists(path):
        fail(f'音频文件不存在：{path}')
    size = os.path.getsize(path)
    if size <= 1000:
        fail(f'音频文件过小：{path} ({size} 字节)')
    dur = mp3_duration(path)
    if dur <= 0:
        fail(f'音频无法解析或时长异常：{path}')
    print(f'[OK] 音频验证通过：{os.path.basename(path)} 大小={size} 时长≈{dur:.0f}s')


def main():
    check_no_overwrite()
    os.makedirs(OUT_DIR, exist_ok=True)

    build_word()

    gen_audio(VOICE_UK, TMP_UK)
    gen_audio(VOICE_US, TMP_US)
    verify_audio(TMP_UK)
    verify_audio(TMP_US)

    if sha256(TMP_UK) == sha256(TMP_US):
        fail('英音和美音哈希相同，判定为异常（疑似复制）')

    os.replace(TMP_UK, MP3_UK)
    os.replace(TMP_US, MP3_US)
    if os.path.exists(AUDIO_TXT):
        os.remove(AUDIO_TXT)

    print('\n===== 完成 =====')
    print(f'周次：第{N}周')
    print(f'主题：{THEME}')
    print(f'难度：{LEVEL}')
    print(f'参考排版：{get_reference_docx()}')
    print(f'Word：{DOCX_PATH}')
    print(f'英音：{MP3_UK}')
    print(f'美音：{MP3_US}')


if __name__ == '__main__':
    main()
