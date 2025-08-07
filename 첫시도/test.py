import plotly
import pandas as pd
import os

import matplotlib.pyplot as plt
import seaborn as sns

import matplotlib.font_manager as fm
import matplotlib

# 1. 사용할 한글 폰트 경로 지정 (예: 맑은고딕)
font_path = "C:/Windows/Fonts/malgun.ttf"  # 또는 NanumGothic.ttf 등
font_prop = fm.FontProperties(fname=font_path)

# 2. matplotlib에 기본 폰트로 설정
matplotlib.rcParams['font.family'] = font_prop.get_name()



data = pd.read_excel('./첫시도/joined_coil_jiwoong.xlsx')

select_data = data[data['wc_desc'] == data['wc_desc'].unique()[0]]

print(data.shape)
print(select_data.shape)

fig, ax = plt.subplots(figsize = (25, 4))
#sns.scatterplot(x = data.index, y = data['wc_id'], ax = ax, hue = data['wc_desc'])
# sns.boxplot(x = data['wc_desc'], y = data['p_thick_mm'], ax = ax, zorder = 1.2)
sns.violinplot(x = data['wc_desc'], y = data['p_thick_mm'], ax = ax, zorder = 1.1)


plt.tight_layout()
plt.show()