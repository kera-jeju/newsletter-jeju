# -*- coding: utf-8 -*-
"""2026년 호(제2호) 설정.

build.py --config config_2026.py 로 불러 쓴다.
대문자 이름만 build.py의 전역값을 덮어쓴다.

※ 원고가 바뀌면 이 파일의 title/subsections만 고치면 된다.
"""

NEWSLETTER_SUBTITLE = "제2호"
PUBLICATION_DATE = "2026년 10월 1일"
COPYRIGHT_YEAR = "2026"
EDITORS = "양은별"

SECTIONS = [
    {
        'id': 'greeting',
        'title': '한국교육학회 제주지회 창립 60주년을 향한 축하말씀',
        'nav_title': '인사말',
        'subtitle': '',
        'icon': '',
        'color': '#1a5c35',
        'file_key': '인사말_고의숙',
        'subsections': []
    },
    {
        'id': 'juron',
        'title': '지나온 60년, 새롭게 열어갈 제주 교육',
        'nav_title': '주론',
        'subtitle': '',
        'icon': '',
        'color': '#1a5c35',
        'file_key': '주론_이인회',
        'subsections': []
    },
    {
        'id': 'siron',
        'title': '시론',
        'nav_title': '시론',
        'subtitle': '시의성 있는 교육 이슈에 대한 소론입니다.',
        'icon': '',
        'color': '#1a5c35',
        'file_key': '',
        'subsections': [
            {'title': '제주지회 60주년 소회와 전망', 'file_key': '시론_김민호'},
            {'title': '시론 (제목 미정)', 'file_key': '시론_연준모'},
        ]
    },
    {
        'id': 'jeju-news',
        'title': '제주교육소식',
        'nav_title': '제주교육소식',
        'subtitle': '제주 지역 교육의 소식과 연구자들의 이야기를 나눕니다.',
        'icon': '',
        'color': '#52b788',
        'file_key': '',
        'subsections': [
            {'title': '성읍마을에서 만난 작은 주인공들',
             'file_key': '제주교육소식_홍지환'},
            {'title': '첼로와 하프 사이에서: 웰니스 음악을 향한 여정',
             'file_key': '제주교육소식_김지원'},
            {'title': '경계를 넘어 런케이션(Learncation)으로',
             'file_key': '제주교육소식_전새미'},
            {'title': '새 교육감에게 바라는 제주 교육의 방향',
             'file_key': '제주교육소식_홍지오'},
            {'title': '말은 하지만, 소통하기 어려운 아이들',
             'file_key': '제주교육소식_석진아'},
            {'title': '지역사회 교육활동의 성과와 제언',
             'file_key': '제주교육소식_김경주'},
            {'title': '박사학위 취득 소감', 'file_key': '제주교육소식_강한호'},
        ]
    },
    {
        'id': 'member',
        'title': '회원 동정',
        'nav_title': '회원 동정',
        'subtitle': '',
        'icon': '',
        'color': '#40916c',
        'file_key': '회원동정',
        'subsections': []
    },
    {
        'id': 'activities',
        'title': '제주지회 활동 소개',
        'nav_title': '활동 소개',
        'subtitle': '',
        'icon': '',
        'color': '#2d6a4f',
        'file_key': '',
        'subsections': [
            {'title': '창립 59주년 기념행사 및 학술발표회',
             'file_key': '활동소개_황현철'},
            {'title': '조천 마을탐방: 항일 정신과 근대교육의 요람을 찾아서',
             'file_key': '활동소개_양유정'},
        ]
    },
]
