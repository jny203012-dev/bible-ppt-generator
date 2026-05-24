let currentFontFamily = '';
let currentBold = false;
let currentSlide = 0;
let slides = [];

console.log("JS 로드됨");

document.getElementById('generateSlides').addEventListener('click', function() {
    const verse = document.getElementById('bibleVerse').value;
    const fontSize = document.getElementById('fontSize').value;
    const bgColor = document.getElementById('bgColor').value;
    const fontColor = document.getElementById('fontColor').value;
    const fontFamily = document.getElementById('fontFamily').value;
    const isBold = document.getElementById('boldToggle').checked;
    const version = document.getElementById('version').value;

    fetch('/preview', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ verse, version })
    })
    .then(res => res.json())
    .then(verses => {
        showSlides(verses, fontSize, bgColor, fontColor, fontFamily, isBold);
    })
    .catch(err => {
        console.error("Fetch error:", err);
        alert("슬라이드를 불러오는 중 오류가 발생했습니다.");
    });
});

function showSlides(verses, fontSize, bgColor, fontColor, fontFamily, isBold) {
    slides = verses;
    currentSlide = 0;

    currentFontFamily = fontFamily;
    currentBold = isBold;

    document.body.innerHTML = `
        <div id="slide" style="
            width:100vw;
            height:100vh;
            display:flex;
            justify-content:center;
            align-items:center;
            background:${bgColor};
            color:${fontColor};
            font-size:${fontSize}px;
            font-family:${fontFamily};
            letter-spacing:0.02em;
            transition:opacity 0.3s;
            position:relative;
            overflow:hidden;
        "></div>
    `;

    renderSlide();
}

function renderSlide() {
    const slide = document.getElementById('slide');

    if (!slide || slides.length === 0) return;

    const verseObj = slides[currentSlide];
    const text = verseObj.text;
    const ref = verseObj.reference;

    slide.style.opacity = 0;

    setTimeout(() => {
        slide.innerHTML = `
            <div style="
                position:absolute;
                top:24px;
                left:32px;
                font-size:0.65em;
                opacity:0.7;
                font-family:${currentFontFamily};
                font-weight:${currentBold ? '700' : '400'};
            ">
                ${ref}
            </div>

            <div style="
                max-width:80vw;
                width:100%;
                text-align:center;
                line-height:1.8;
                padding:60px;
                box-sizing:border-box;
                font-family:${currentFontFamily};
                font-weight:${currentBold ? '700' : '400'};
                word-break:keep-all;
                overflow-wrap:break-word;
                white-space:normal;
            ">
                <div>
                    ${text}
                </div>

                <div style="
                    font-size:0.6em;
                    opacity:0.6;
                    margin-top:24px;
                ">
                    ${currentSlide + 1} / ${slides.length}
                </div>
            </div>
        `;

        slide.style.opacity = 1;
    }, 100);
}

document.addEventListener('click', (e) => {
    if (slides.length === 0) return;

    const half = window.innerWidth / 2;

    if (e.clientX > half) {
        currentSlide++;
    } else {
        currentSlide--;
    }

    if (currentSlide < 0) currentSlide = 0;
    if (currentSlide >= slides.length) currentSlide = slides.length - 1;

    renderSlide();
});

document.addEventListener('keydown', (e) => {
    if (slides.length === 0) return;

    if (e.key === 'ArrowRight') {
        currentSlide++;
    } else if (e.key === 'ArrowLeft') {
        currentSlide--;
    } else if (e.key === 'Escape') {
        location.reload();
    }

    if (currentSlide < 0) currentSlide = 0;
    if (currentSlide >= slides.length) currentSlide = slides.length - 1;

    renderSlide();
});

document.getElementById('downloadPPT').addEventListener('click', function() {
    const button = document.getElementById('downloadPPT');
    const loading = document.getElementById('loadingText');

    const verse = document.getElementById('bibleVerse').value;
    const fontSize = document.getElementById('fontSize').value;
    const bgColor = document.getElementById('bgColor').value;
    const fontColor = document.getElementById('fontColor').value;
    const fontFamily = document.getElementById('fontFamily').value;
    const isBold = document.getElementById('boldToggle').checked;
    const version = document.getElementById('version').value;

    if (!verse.trim()) {
        alert("성경 구절을 입력해주세요.");
        return;
    }

    button.disabled = true;
    button.innerText = '생성 중...';
    loading.style.display = 'block';

    const params = new URLSearchParams({
        verse,
        fontSize,
        bgColor,
        fontColor,
        fontFamily,
        isBold,
        version,
        t: Date.now()
    });

// 기존 iframe 제거
const oldFrame = document.getElementById('downloadFrame');
if (oldFrame) {
    oldFrame.remove();
}

// 새 iframe 생성
const iframe = document.createElement('iframe');
iframe.style.display = 'none';
iframe.id = 'downloadFrame';
iframe.src = `/generate?${params.toString()}`;

document.body.appendChild(iframe);

// 다운로드 완료 감지가 불안정해서 일정 시간 후 UI만 복구
setTimeout(() => {
    button.disabled = false;
    button.innerText = 'PPT 저장';
    loading.style.display = 'none';

    const frame = document.getElementById('downloadFrame');
    if (frame) {
        frame.remove();
    }
}, 12000);