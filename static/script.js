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

document.getElementById('downloadPPT').addEventListener('click', async function() {
    const button = document.getElementById('downloadPPT');
    const loading = document.getElementById('loadingText');

    button.disabled = true;
    button.innerText = '생성 중...';
    loading.style.display = 'block';
    
    const verse = document.getElementById('bibleVerse').value;
    const fontSize = document.getElementById('fontSize').value;
    const bgColor = document.getElementById('bgColor').value;
    const fontColor = document.getElementById('fontColor').value;
    const fontFamily = document.getElementById('fontFamily').value;
    const isBold = document.getElementById('boldToggle').checked;
    const version = document.getElementById('version').value;

    if (!verse.trim()) {
        alert("성경 구절을 입력해주세요. 예: 겔40:1-4");
        return;
    }

    try {
        const response = await fetch('/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                verse,
                fontSize,
                bgColor,
                fontColor,
                fontFamily,
                isBold,
                version
            })
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error("서버 오류:", errorText);
            alert("PPT 저장 오류:\n" + errorText);
            return;
        }

        // 파일명 추출
        const disposition = response.headers.get('Content-Disposition');
        let filename = 'bible_slides.pptx';

        if (disposition) {
            // filename*=UTF-8''... 형식 우선 처리
            const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/);
            if (utf8Match && utf8Match[1]) {
                filename = decodeURIComponent(utf8Match[1]);
            } else {
                // 일반 filename="..." 형식 처리
                const normalMatch = disposition.match(/filename="?([^"]+)"?/);
                if (normalMatch && normalMatch[1]) {
                    filename = normalMatch[1];
                }
            }
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = filename;

        document.body.appendChild(a);
        a.click();

        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        button.disabled = false;
        button.innerText = 'PPT 저장';
        loading.style.display = 'none';

    } catch (err) {
        console.error("PPT 저장 실패:", err);
        alert("PPT 저장 중 오류가 발생했습니다.");
    } finally {
        button.disabled = false;
        button.innerText = 'PPT 저장';
        loading.style.display = 'none';
    }
});
