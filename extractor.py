import pandas as pd
import json

# 假设你的JSON数据存储在一个名为'resumes.json'的文件中
with open('resumes.json', 'r', encoding='utf-8') as json_file:
    # 加载JSON数据
    data = json.load(json_file)

# 初始化一个空的DataFrame
df = pd.DataFrame()

# 遍历JSON数据中的每个简历
for key, resume in data.items():
    # 提取基本信息
    basic_info = {
        '序号': key,  # 序号作为第一列
        '姓名': resume.get('姓名', '无姓名'),
        '电话': resume.get('电话', '无电话'),
        '政治面貌': resume.get('政治面貌', '无政治面貌'),
        '籍贯': resume.get('籍贯', '无籍贯'),
        '出生年月': resume.get('出生年月', '无出生年月'),
        '落户市县': resume.get('落户市县', '无落户市县'),
        # 获取最高学历
        '最高学历': max((edu.get('学位', '无学位') for edu in resume.get('教育经历', [])), default='无学位') if '教育经历' in resume else '无学位',
        # 项目经历
        '项目经历': '\n'.join(f"{item.get('项目名称', '无项目名称')} - {item.get('项目责任', '无项目责任')} ({item.get('项目时间', '无项目时间')})" for item in resume.get('项目经历', [])) if '项目经历' in resume else '',
        # 工作经历
        '工作经历': '\n'.join(f"{item.get('工作单位', '无工作单位')} - {item.get('职务', '无职务')} ({item.get('工作时间', '无工作时间')})" for item in resume.get('工作经历', [])) if '工作经历' in resume else ''
    }
    
    # 创建一个包含当前简历信息的DataFrame
    resume_df = pd.DataFrame([basic_info])
    
    # 使用pd.concat()将当前简历的DataFrame与之前的DataFrame合并
    df = pd.concat([df, resume_df], ignore_index=True)

# 保存DataFrame到Excel文件
df.to_excel('resumes.xlsx', index=False)