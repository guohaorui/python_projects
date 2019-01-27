#!/usr/bin/env python 
# -*- coding:utf-8 -*-
import requests,docx
import os
from lxml import etree
from innovation import downloadFile
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.92 Safari/537.36'
}
BASE_DOMAIN = 'https://www.superlearn.com/'
Listening_url = 'https://www.superlearn.com/practice/base/listen/topic_38_154.shtml'
# considering that my account maybe shut up, switch it to 00000, also work, it is a bug
BASE_Listening_text_url = 'https://www.superlearn.com/quiz-bank/ExaminationNew/loadIntensiveListening?qId={}&userId=00000'
BASE_Listening_material_url = 'https://image.superlearn.com/upload/question/{}.mp3'
BASE_Listening_questions_url = 'https://www.superlearn.com/quiz-bank/toefl/question/details/{}?qId={}'

class Material(object):
    def __init__(self):
        self.title = None
        self.questions_set = []
        self.listening_cn_text = []
        self.listening_en_text = []
        self.listening_simple_en_text = []
        self.normal = True

# 1. 获取听力conversation页
def get_home_page(Listening_url):
    print('=' * 30)
    print('geting the conversation page')
    response = requests.get(Listening_url, headers=HEADERS)
    with open('superlearn.html', 'w', encoding='utf-8') as fp:
        fp.write(response.content.decode('utf-8'))
    text = response.content.decode('utf-8')
    html = etree.HTML(text)
    print('=' * 30)
    return html

# 2. 解析各个文章的精听听写和题目
# analyse the listening text url
# qId is the key to get the text of Listening
def analyse_listenings(html):
    print('=' * 30)
    print('analyse the listening text url')
    detail_text_urls = html.xpath(
        "//div[@class='Tn-llbox']//ul//li[@class='Tn-wid210 Tn-ll-right Tn-tright']//a[@target='_blank']/@href")
    qIds = []
    Listening_text_urls = []
    for detail_text_url in detail_text_urls:
        qId = detail_text_url[detail_text_url.find('qId') + 4:detail_text_url.find('&')]
        qIds.append(qId)
        Listening_text_url = BASE_Listening_text_url.format(qId)
        Listening_text_urls.append(Listening_text_url)

    # analyse the listening material url
    print('analyse the listening material url')
    titles = []
    Listening_material_urls = []
    detail_titles = html.xpath("//div[@class='Tn-wid335']/h3/a/text()")
    for detail_title in detail_titles:
        titles.append(detail_title)
        num = detail_title[-1]
        if num == '1':
            material_num = detail_title[4:7] + detail_title[-1]
        elif num == '2':
            material_num = detail_title[4:7] + '4'
        Listening_material_url = BASE_Listening_material_url.format(material_num)
        Listening_material_urls.append(Listening_material_url)
    # print(Listening_material_urls)

    # analyse the listening material questions
    print('analyse the listening material questions')
    Listening_questions_urls = []
    for qId in qIds:
        Listening_questions_url = BASE_Listening_questions_url.format(qId, qId)
        Listening_questions_urls.append(Listening_questions_url)
    # print(Listening_questions_urls)
    print('=' * 30)
    return Listening_text_urls,Listening_material_urls,Listening_questions_urls,titles,detail_titles

# 3. get listening text, material, questions
def get_listenings(Listening_text_urls, Listening_material_urls, Listening_questions_urls, titles, detail_titles,
                   materials):
    # listening text
    print('=' * 30)
    print('get listening text')
    for idx, Listening_text_url in enumerate(Listening_text_urls):
        response = requests.get(Listening_text_url, headers=HEADERS)
        texts = response.content.decode().split(',')
        cntexts = []
        entexts = []
        for text in texts:
            if text.startswith('"cntext"'):
                cntexts.append(text)
            if text.startswith('"entext"'):
                entexts.append(text)
        materials[idx].listening_cn_text = cntexts
        materials[idx].listening_en_text = entexts
        materials[idx].title = detail_titles[idx]
    #    break
    # listening material
    print('get listening material')
    for idx, Listening_material_url in enumerate(Listening_material_urls):
        name = 'C:\\Users\\Administrator\\Desktop\\英语学习\\听力练习\\Conversation\\听力音频\\{}_{}.mp3'.format(str(idx + 1),
                                                                                                     titles[idx])
        if os.path.exists(name):
            pass
        else:
            downloadFile(name, url=Listening_material_url)
    # listening material questions
    print('get listening questions')
    count = 0
    for idx1, Listening_questions_url in enumerate(Listening_questions_urls):
        response = requests.get(Listening_questions_url, headers=HEADERS)
        text = response.content.decode('utf-8')
        html = etree.HTML(text)
        questions = html.xpath("//div[@class='nzkStem']/p/text()")
        questions = questions[0:5]
        answers_per_question = html.xpath("//ul[@class='nzkChoice']")
        if len(answers_per_question) < 5:
            count = count + 1
            materials[idx].normal = False
            continue
        per_question = {}
        questions_set = []
        for idx2, item in enumerate(questions):
            answers = answers_per_question[idx2].xpath(".//li//span/text()")
            per_question['question'] = item
            per_question['answer1'] = answers[0]
            per_question['answer2'] = answers[1]
            per_question['answer3'] = answers[2]
            per_question['answer4'] = answers[3]
            questions_set.append(per_question)
            per_question = {}

        materials[idx1].questions_set = questions_set
        questions_set = {}
        simple_en_text = html.xpath("//div[@id='article']//p/text()")
        materials[idx1].listening_simple_en_text = simple_en_text
    print('get listening information done')
    print('=' * 30)
def write_file(materials):
    print('=' * 30)
    print('writing file to docx ')
    people = ['Narrator', 'Student', 'Professor', 'W', 'M', 'STUDENT', 'PROFESSOR',
              'P','S']
    for idx1,material in enumerate(materials):
        doc = docx.Document()
        doc.styles['Normal'].font.name = u'微软雅黑'
        if material.normal:
            doc.add_paragraph(material.title)
            # load en_text
            for text in material.listening_simple_en_text:
                doc.add_paragraph(text)
            doc.add_paragraph('\n\n\n')
            # load questions
            doc.add_paragraph('The Questions:')
            for text in material.questions_set:
                text['question'] = text['question'].replace('\n','')
                doc.add_paragraph(text['question'])
                doc.add_paragraph(text['answer1'])
                doc.add_paragraph(text['answer2'])
                doc.add_paragraph(text['answer3'])
                doc.add_paragraph(text['answer4'])
                doc.add_paragraph('\n')
        else:
            doc.add_paragraph(material.title)
            doc.add_paragraph('the question type is not normal, please check the web')
        doc.save('C:\\Users\\Administrator\\Desktop\\英语学习\\听力练习\\Conversation\\听力原文\\{}_{}.docx'.format(str(idx1 + 1), material.title))
    print('writing successfully')
    print('=' * 30)

if __name__ == '__main__':
    materials = []
    for i in range(1, 60):
        material = Material()
        materials.append(material)
    html = get_home_page(Listening_url)
    Listening_text_urls, Listening_material_urls, Listening_questions_urls, titles, detail_titles=analyse_listenings(html)
    get_listenings(Listening_text_urls,Listening_material_urls,Listening_questions_urls,titles,detail_titles,materials)
    write_file(materials)