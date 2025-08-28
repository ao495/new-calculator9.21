document.addEventListener('DOMContentLoaded', () => {
    // --- DOM ELEMENTS ---
    const displayResult = document.querySelector('.result');
    const displayHistory = document.querySelector('.history');
    const normalCalculatorPanel = document.getElementById('normal-calculator-panel');
    const dateCalculatorPanel = document.getElementById('date-calculator-panel');
    const timeCalculatorPanel = document.getElementById('time-calculator-panel');
    const modeSwitcher = document.querySelector('.mode-switcher');

    // --- MODE SWITCHING ---
    modeSwitcher.addEventListener('click', (e) => {
        if (!e.target.matches('.mode-button')) return;

        modeSwitcher.querySelector('.active').classList.remove('active');
        e.target.classList.add('active');

        const mode = e.target.dataset.mode;
        if (mode === 'normal') {
            normalCalculatorPanel.style.display = 'block';
            dateCalculatorPanel.style.display = 'none';
            timeCalculatorPanel.style.display = 'none';
            displayHistory.style.display = 'block';
            updateDisplay();
        } else if (mode === 'date') {
            normalCalculatorPanel.style.display = 'none';
            dateCalculatorPanel.style.display = 'block';
            timeCalculatorPanel.style.display = 'none';
            displayHistory.style.display = 'none';
            displayResult.textContent = '日付計算';
        } else if (mode === 'time') {
            normalCalculatorPanel.style.display = 'none';
            dateCalculatorPanel.style.display = 'none';
            timeCalculatorPanel.style.display = 'block';
            displayHistory.style.display = 'none';
            displayResult.textContent = '時間計算';
        }
    });

    // --- NORMAL CALCULATOR LOGIC ---
    const buttons = normalCalculatorPanel.querySelector('.buttons');
    const allButtons = normalCalculatorPanel.querySelectorAll('.button');
    const TAX_RATE = 0.10;

    let currentInput = '0';
    let operator = null;
    let previousInput = null;

    buttons.addEventListener('click', (e) => {
        if (!e.target.matches('button')) return;
        const button = e.target;
        button.classList.add('button--active');
        setTimeout(() => { button.classList.remove('button--active'); }, 150);
        handleAction(button.textContent);
        updateDisplay();
    });

    document.addEventListener('keydown', (e) => {
        if (normalCalculatorPanel.style.display === 'none') return;
        if (e.repeat) return;
        let key = e.key;
        let action;

        if (!isNaN(key) || key === '.') { action = key; }
        else if (['+', '-'].includes(key)) { action = key; }
        else if (key === '*') { action = '×'; }
        else if (key === '/') { action = '÷'; }
        else if (key.toLowerCase() === 's') { action = '+/-'; }
        else if (key === 'Enter' || key === '=') { action = '='; }
        else if (key === 'Escape' || key === 'Delete') { action = 'AC'; }
        else if (key === 'Backspace') { action = 'DEL'; }
        else { return; }

        e.preventDefault();
        const button = [...allButtons].find(b => b.textContent === action);
        if (button) {
            button.classList.add('button--active');
            setTimeout(() => { button.classList.remove('button--active'); }, 150);
        }
        handleAction(action);
        updateDisplay();
    });

    function handleAction(action) {
        if (!isNaN(action) || action === '.') { handleNumber(action); }
        else if (['+', '-', '×', '÷'].includes(action)) { handleOperator(action); }
        else { handleSpecial(action); }
    }

    function handleNumber(number) {
        if (currentInput === '0' && number !== '.') { currentInput = number; }
        else if (number === '.' && currentInput.includes('.')) { return; }
        else { currentInput += number; }
    }

    function handleOperator(op) {
        if (operator !== null) { calculate(); }
        previousInput = currentInput;
        currentInput = '0';
        operator = op;
        displayHistory.textContent = `${previousInput} ${operator}`;
    }

    function calculate() {
        if (previousInput === null || operator === null) return;
        let result;
        const prev = parseFloat(previousInput);
        const curr = parseFloat(currentInput);
        switch (operator) {
            case '+': result = prev + curr; break;
            case '-': result = prev - curr; break;
            case '×': result = prev * curr; break;
            case '÷':
                if (curr === 0) { alert('0で割ることはできません'); return; }
                result = prev / curr;
                break;
        }
        currentInput = result.toString();
        operator = null;
        previousInput = null;
        displayHistory.textContent = '';
    }

    function handleSpecial(action) {
        const value = parseFloat(currentInput);
        switch (action) {
            case 'AC':
                currentInput = '0'; operator = null; previousInput = null; displayHistory.textContent = ''; break;
            case 'DEL':
                currentInput = currentInput.slice(0, -1) || '0'; break;
            case '+/-':
                if (currentInput === '0') return;
                currentInput = (value * -1).toString(); break;
            case '税込':
                currentInput = (value * (1 + TAX_RATE)).toString(); break;
            case '税抜':
                currentInput = (value / (1 + TAX_RATE)).toString(); break;
            case '=': calculate(); break;
        }
    }

    function formatNumberWithCommas(numberString) {
        const isNegative = numberString.startsWith('-');
        let num = isNegative ? numberString.substring(1) : numberString;
        const parts = num.split('.');
        let integerPart = parts[0];
        const decimalPart = parts.length > 1 ? '.' + parts[1] : '';
        integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        return (isNegative ? '-' : '') + integerPart + decimalPart;
    }

    function updateDisplay() { displayResult.textContent = formatNumberWithCommas(currentInput); }

    // --- DATE CALCULATOR LOGIC ---
    const baseDateInput = document.getElementById('base-date');
    const daysToAddInput = document.getElementById('days-to-add');
    const addDaysButton = document.getElementById('add-days');
    const subtractDaysButton = document.getElementById('subtract-days');
    const calculatedDateSpan = document.getElementById('calculated-date');
    const countStartDayCheckbox = document.getElementById('count-start-day');

    function performDateCalculation(operation) {
        const baseDateValue = baseDateInput.value;
        let days = parseInt(daysToAddInput.value, 10);

        if (!baseDateValue) {
            calculatedDateSpan.textContent = "基準日を選択してください";
            return;
        }
        if (isNaN(days)) {
            calculatedDateSpan.textContent = "日数を入力してください";
            return;
        }

        if (countStartDayCheckbox.checked) {
            if (operation === 'add') {
                days -= 1;
            } else if (operation === 'subtract') {
                days += 1;
            }
        }

        const baseDate = new Date(baseDateValue);
        baseDate.setMinutes(baseDate.getMinutes() + baseDate.getTimezoneOffset());

        if (operation === 'add') {
            baseDate.setDate(baseDate.getDate() + days);
        } else if (operation === 'subtract') {
            baseDate.setDate(baseDate.getDate() - days);
        }

        const year = baseDate.getFullYear();
        const month = (baseDate.getMonth() + 1).toString().padStart(2, '0');
        const day = baseDate.getDate().toString().padStart(2, '0');
        const formattedDate = `${year}-${month}-${day}`;

        calculatedDateSpan.textContent = formattedDate;
        displayResult.textContent = formattedDate;
    }

    addDaysButton.addEventListener('click', () => performDateCalculation('add'));
    subtractDaysButton.addEventListener('click', () => performDateCalculation('subtract'));

    // --- TIME CALCULATOR LOGIC ---
    const startHoursInput = document.getElementById('start-hours');
    const startMinutesInput = document.getElementById('start-minutes');
    const endHoursInput = document.getElementById('end-hours');
    const endMinutesInput = document.getElementById('end-minutes');
    const calculateTimeDiffButton = document.getElementById('calculate-time-diff');
    const timeDiffResultSpan = document.getElementById('time-diff-result');

    function timeToMinutes(h, m) {
        return (h * 60) + m;
    }

    function minutesToTime(totalMinutes) {
        const h = Math.floor(totalMinutes / 60);
        const m = totalMinutes % 60;
        return { h, m };
    }

    function formatTime(h, m) {
        const pad = (num) => num.toString().padStart(2, '0');
        return `${pad(h)}:${pad(m)}`;
    }

    calculateTimeDiffButton.addEventListener('click', () => {
        const startHours = parseInt(startHoursInput.value) || 0;
        const startMinutes = parseInt(startMinutesInput.value) || 0;
        const endHours = parseInt(endHoursInput.value) || 0;
        const endMinutes = parseInt(endMinutesInput.value) || 0;

        const startTimeInMinutes = timeToMinutes(startHours, startMinutes);
        let endTimeInMinutes = timeToMinutes(endHours, endMinutes);

        if (endTimeInMinutes < startTimeInMinutes) {
            endTimeInMinutes += (24 * 60); 
        }

        let diffMinutes = endTimeInMinutes - startTimeInMinutes;

        if (diffMinutes < 0) {
            timeDiffResultSpan.textContent = "エラー";
            displayResult.textContent = "エラー";
            return;
        }

        const diffTime = minutesToTime(diffMinutes);
        const formattedResult = formatTime(diffTime.h, diffTime.m);

        timeDiffResultSpan.textContent = formattedResult;
        displayResult.textContent = formattedResult;
    });
});

