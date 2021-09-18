import base64
import os

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from pandas.core.frame import DataFrame


def get_binary_file_downloader_html(bin_file, file_label='File'):
    with open(bin_file, 'rb') as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}">Download</a>'
    return href

def df_to_text(df:DataFrame, separation_letter:str, space1:str, space2:str, delete_zero:bool):
    text_list = []
    for row in df.itertuples():
        if '■' in row.検査項目:
            continue
        elif row.検査項目 == '':
            text_item = '\n\n'
        elif row.検査項目 != '' and row.数値 == '':
            continue
        else:
            if delete_zero and row.数値.endswith('.0'):
                num = row.数値.replace('.0', '')
            else:
                num = row.数値
            unit = '%)' if row.検査項目 == 'Lymphocyte' else row.単位
            text_item = f'{row.検査項目}{space1}{num}{space2}{unit}'
        text_list.append(text_item)
    while text_list[-1] == '\n\n':
        del text_list[-1]
    text = separation_letter.join(text_list)
    text = text.replace(', Seg', ' (Seg').replace('、Seg', ' (Seg')
    text = text.replace(', \n\n, ', '\n\n').replace('、\n\n、', '\n\n')
    while '\n\n\n\n' in text:
        text = text.replace('\n\n\n\n', '\n\n')
    return text

st.set_page_config(
    page_title='検査結果 文字起こしツール',
    initial_sidebar_state='expanded'
)


st.title('検査結果 文字起こしツール')

# サイドバー
st.sidebar.write('# オプション')

st.sidebar.write('### 区切りに使用する文字')
separation_letter = st.sidebar.selectbox('選択してください', (',', '、'))
separation_letter = separation_letter.replace(',', ', ')

st.sidebar.write('### スペースの挿入位置')
space_insertion = st.sidebar.selectbox('選択してください', ('赤血球420万（なし）', '赤血球 420万（項目と数値の間）', '赤血球 420 万（数値の前後）'))
space1, space2 = '', ''
if space_insertion == '赤血球420万（なし）':
    pass
elif space_insertion == '赤血球 420万（項目と数値の間）':
    space1 = ' '
else:
    space1, space2 = ' ', ' '

st.sidebar.write('### 小数点以下の処理')
which_zero = st.sidebar.selectbox('選択してください', ('420（「.0」を非表示）', '420.0（「.0」を表示）'))
delete_zero = False
if which_zero == '420（「.0」を非表示）':
    delete_zero = True


# ツール
try:
    st.text('')
    st.write('### ▼ テンプレート（Excelファイル）をダウンロードし、数値を入力')
    st.text('※事前のダウンロードをおすすめします※')
    st.markdown(get_binary_file_downloader_html('template.xlsx'), unsafe_allow_html=True)
    st.text('')
    st.write('### ▼ 数値を入力したExcelファイルをアップロード')
    uploaded_file = st.file_uploader(
        'ファイルを選択',
        type='xlsx'
    )
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        df = df.astype(str)
        df = df.replace('nan', '')
        if st.checkbox('ファイルの中身を表示'):
            df_show = df
            st.table(df_show)
        st.text('')
        st.write('### ▼ 文字起こし（数値が空白の部分はスキップします）')
        if st.button('文字起こし'):
            st.write('#### 【出力結果】')
            st.write('右上のアイコンからコピーできます')
            st.code(df_to_text(df, separation_letter, space1, space2, delete_zero), language='text')
except:
    st.error('エラーが発生しました')


# できること
st.text('')
st.write('***')
st.write('# できること')
st.write("""
臨床検査の結果をExcelの表に打ち込んだものを、レポートや論文にコピペできる形式でテキストとして出力します。\n
病院内等でインターネットに接続できない場合は、あらかじめExcelテンプレートをダウンロードしておくことをおすすめします。
""")
st.text('')
st.image('images/screenshot2.png')
st.text('')
st.write('⬇ 変換')
st.image('images/screenshot7.png')


# 使い方
st.text('')
st.write('***')
st.write('# 使い方')
st.markdown('### ①テンプレート（Excelファイル）をダウンロードする')
st.markdown(get_binary_file_downloader_html('template.xlsx'), unsafe_allow_html=True)
st.write('ファイルの中身はこのようになっています。')
st.image('images/screenshot1.png')
st.text('')
st.text('')
st.text('')
st.write("""
#### ↓↓↓ここからオフライン作業可能↓↓↓
### ②テンプレートに数値を入力する
数値を入力してください。
項目や単位の追加・変更にも対応しています。
""")
st.image('images/screenshot2.png')
st.write('#### ↑↑↑オフライン作業ここまで↑↑↑')
st.text('')
st.text('')
st.text('')
st.write('### ③編集済みのExcelファイルをアップロードする')
st.write('ここにアップロードしたファイルが保存されることはありません。')
st.image('images/screenshot3.png')
st.text('')
st.text('')
st.write('### （④ファイルの中身を確認する）')
st.write('チェックを入れるとExcelファイルの中身を確認できます。チェックを外すと閉じます。')
st.image('images/screenshot4.png')
st.text('')
st.text('')
st.write('### ⑤文字起こしのオプションを設定する')
st.write("""
文字起こしする際のオプションを設定します。\n
オプションが表示されない場合は、画面左上の「＞」を押してください。
- 項目の区切りに「,」「、」どちらを使用するか
- 項目、数値、単位の間にスペース（空白）を挿入するか
- 数値末尾の「.0」を表示させるか\n
を選択してください。
""")
st.image('images/screenshot5.png')
st.text('')
st.text('')
st.write('### ⑥文字起こしを行う')
st.write("""
文字起こしを実行します。
出力された文字は、テキストボックス右上に表示されるアイコンをクリックすることで、クリップボードにコピーされます。\n
Word等に「貼り付け（ペースト）」してご使用ください。\n
出力をした後でも、オプションを変更して再出力することもできます。
""")
st.image('images/screenshot6.png')


# コメント
st.write('***')
st.write("""**Made by Cyrus ( <a href="https://cyrus.tokyo" target="_blank">HP</a> / <a href="https://twitter.com/cyrus_twi" target="_blank">Twitter</a> )**""", unsafe_allow_html=True)
st.text('')
st.write("""ご意見・エラー報告・改善案 等は <a href="https://cyrus.tokyo/contact" target="_blank">こちら</a> または Twitter にお寄せください。""", unsafe_allow_html=True)
st.write('※ 特にテンプレートの記載内容についてのご意見を募集しております')
components.html(
    """
    <a href="https://twitter.com/share?ref_src=twsrc%5Etfw" class="twitter-share-button" data-size="large" data-text="【検査結果 文字起こしツール】" data-url="https://share.streamlit.io/cyrus-git/medicaltest-to-text/main/app.py" data-via="cyrus_twi" data-hashtags="streamlit" data-lang="ja" data-show-count="false">Tweet</a><script async src="https://platform.twitter.com/widgets.js" charset="utf-8"></script>
    """
)
st.text('sponsored link')
st.markdown("""
    <a href="https://www.amazon.co.jp/studentawgateway/?tag=gondo238-22" target="_blank">
        <img src="https://images-fe.ssl-images-amazon.com/images/G/09/JP-hq/2021/img/Prime/XCM_Manual_1353312_1791081_3998434_JP_assoc_728x90_ja_JP.jpg" />
    </a>""",
    unsafe_allow_html=True
)


