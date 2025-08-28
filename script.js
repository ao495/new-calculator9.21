document.addEventListener('DOMContentLoaded', () => {
    // --- DOM Elements ---
    const timerDisplay = document.getElementById('timer-display');
    const startBtn = document.getElementById('start-btn');
    const stopBtn = document.getElementById('stop-btn');
    const resetBtn = document.getElementById('reset-btn');
    
    const minutesInput = document.getElementById('minutes-input');
    const secondsInput = document.getElementById('seconds-input');
    
    const settingsBox = document.getElementById('settings-box');
    const stopwatchModeBtn = document.getElementById('stopwatch-mode-btn');
    const countdownModeBtn = document.getElementById('countdown-mode-btn');

    // --- State Variables ---
    let startTime;
    let timerInterval;
    let isRunning = false;
    let currentMode = 'stopwatch'; // 'stopwatch' or 'countdown'

    let stopwatchElapsedTime = 0;
    let countdownRemainingTime = 0;

    // --- Core Functions ---
    function formatTime(ms) {
        if (ms < 0) ms = 0;
        const totalSeconds = Math.floor(ms / 1000);
        const minutes = String(Math.floor(totalSeconds / 60)).padStart(2, '0');
        const seconds = String(totalSeconds % 60).padStart(2, '0');
        const milliseconds = String(ms % 1000).padStart(3, '0');
        return `${minutes}:${seconds}.${milliseconds}`;
    }

    function switchMode(newMode) {
        if (isRunning) stopTimer();
        resetTimer(false); // Soft reset without stopping again

        currentMode = newMode;

        if (newMode === 'stopwatch') {
            settingsBox.classList.add('hidden');
            stopwatchModeBtn.classList.add('active');
            countdownModeBtn.classList.remove('active');
        } else {
            settingsBox.classList.remove('hidden');
            stopwatchModeBtn.classList.remove('active');
            countdownModeBtn.classList.add('active');
        }
    }

    // --- Timer Logic ---
    function startTimer() {
        if (isRunning) return;
        
        if (currentMode === 'countdown') {
            const inputMinutes = parseInt(minutesInput.value) || 0;
            const inputSeconds = parseInt(secondsInput.value) || 0;
            const totalInputSeconds = inputMinutes * 60 + inputSeconds;

            if (countdownRemainingTime <= 0 && totalInputSeconds > 0) {
                countdownRemainingTime = totalInputSeconds * 1000;
            }

            if (countdownRemainingTime <= 0) {
                alert('カウントダウンの時間を設定してください。');
                return;
            }
        }

        isRunning = true;
        startTime = Date.now();
        timerInterval = setInterval(currentMode === 'stopwatch' ? updateStopwatch : updateCountdown, 10);
    }

    function stopTimer() {
        if (!isRunning) return;
        isRunning = false;
        clearInterval(timerInterval);

        const timePassedSinceStart = Date.now() - startTime;

        if (currentMode === 'countdown') {
            countdownRemainingTime -= timePassedSinceStart;
        } else {
            stopwatchElapsedTime += timePassedSinceStart;
        }
    }

    function resetTimer(hardReset = true) {
        if (hardReset && isRunning) {
            stopTimer();
        }

        stopwatchElapsedTime = 0;
        countdownRemainingTime = 0;
        minutesInput.value = "";
        secondsInput.value = "";
        timerDisplay.textContent = formatTime(0);
    }

    // --- Update Functions ---
    function updateStopwatch() {
        const timePassedSinceStart = Date.now() - startTime;
        timerDisplay.textContent = formatTime(stopwatchElapsedTime + timePassedSinceStart);
    }

    function updateCountdown() {
        const timePassedSinceStart = Date.now() - startTime;
        const currentRemaining = countdownRemainingTime - timePassedSinceStart;

        if (currentRemaining <= 0) {
            resetTimer();
            timerDisplay.textContent = formatTime(0);
            alert("時間です！");
            switchMode('stopwatch'); // Switch back to stopwatch mode
            return;
        }
        timerDisplay.textContent = formatTime(currentRemaining);
    }

    // --- Event Listeners ---
    startBtn.addEventListener('click', startTimer);
    stopBtn.addEventListener('click', stopTimer);
    resetBtn.addEventListener('click', () => resetTimer(true));

    stopwatchModeBtn.addEventListener('click', () => switchMode('stopwatch'));
    countdownModeBtn.addEventListener('click', () => switchMode('countdown'));

    document.addEventListener('keydown', (e) => {
        // Ignore if typing in an input field
        if (e.target.tagName === 'INPUT') return;

        if (e.key === 'Shift') {
            e.preventDefault(); // Prevent any default browser action
            if (isRunning) {
                stopTimer();
            } else {
                startTimer();
            }
        }
    });
});

