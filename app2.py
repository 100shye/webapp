from flask import Flask, request, render_template_string
import difflib
import json
import os

app = Flask(__name__)

# 예시 문서 데이터 (문서 ID와 내용)
documents = {
    1: "Flask는 Python으로 작성된 마이크로 웹 프레임워크입니다.",
    2: "역색인은 정보 검색 시스템에서 자주 사용되는 기법입니다.",
    3: "검색 엔진은 다양한 알고리즘을 사용하여 문서를 추천합니다.",
    4: "유사도 매칭은 사용자의 오타나 변형된 표현을 처리하는데 유용합니다.",
}

# 영구 저장을 위한 역색인 파일 경로
index_file = 'inverted_index.json'

# 역색인 로드 또는 생성
if os.path.exists(index_file):
    with open(index_file, 'r', encoding='utf-8') as f:
        # JSON은 리스트로 저장하므로, 값을 set으로 변환
        inverted_index = {key: set(value) for key, value in json.load(f).items()}
else:
    inverted_index = {}
    for doc_id, content in documents.items():
        for word in content.split():
            # 소문자로 변환하고 구두점 제거
            word = word.lower().strip('.,!?')
            if word in inverted_index:
                inverted_index[word].add(doc_id)
            else:
                inverted_index[word] = {doc_id}
    # JSON 파일에 저장하기 위해 set을 list로 변환
    inverted_index_to_save = {key: list(value) for key, value in inverted_index.items()}
    with open(index_file, 'w', encoding='utf-8') as f:
        json.dump(inverted_index_to_save, f, ensure_ascii=False, indent=2)

@app.route('/', methods=['GET', 'POST'])
def search():
    results = {}
    query = ""
    if request.method == 'POST':
        query = request.form.get('query', '').lower()
        query_words = query.split()
        doc_ids = set()
        # 각 쿼리 단어에 대해 역색인 검색 및 유사 단어 추천
        for q in query_words:
            if q in inverted_index:
                doc_ids.update(inverted_index[q])
            else:
                # 유사한 단어 찾기 (최대 1개, cutoff 조정 가능)
                close = difflib.get_close_matches(q, inverted_index.keys(), n=1, cutoff=0.6)
                if close:
                    doc_ids.update(inverted_index[close[0]])
        # 결과 문서 내용 가져오기
        results = {doc_id: documents[doc_id] for doc_id in doc_ids}

    # 간단한 HTML 템플릿 (실제 서비스에서는 별도의 템플릿 파일 권장)
    html = """
    <!doctype html>
    <html>
    <head>
        <title>역색인 검색</title>
    </head>
    <body>
        <h1>문서 검색</h1>
        <form method="post">
            <input type="text" name="query" value="{{ query }}" placeholder="검색어를 입력하세요">
            <input type="submit" value="검색">
        </form>
        <h2>검색 결과</h2>
        {% if results %}
            <ul>
            {% for doc_id, content in results.items() %}
                <li><strong>문서 {{ doc_id }}</strong>: {{ content }}</li>
            {% endfor %}
            </ul>
        {% else %}
            <p>검색 결과가 없습니다.</p>
        {% endif %}
    </body>
    </html>
    """
    return render_template_string(html, results=results, query=query)

if __name__ == '__main__':
    app.run(debug=True)
