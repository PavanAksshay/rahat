/**
 * RAHAT — MAKING SPACE
 * Inaugural Presentation Controller (9-Slide Final Master Edition)
 */

class PresentationDeck {
  constructor() {
    this.slides = Array.from(document.querySelectorAll('.slide'));
    this.totalSlides = this.slides.length;
    this.currentIndex = 0;
    
    // Ambient Audio Engine
    this.audioCtx = null;
    this.isAudioPlaying = false;
    this.audioGainNode = null;
    this.oscillators = [];
    
    // UI Elements
    this.slideNumEl = document.getElementById('slide-num');
    this.totalSlidesEl = document.getElementById('total-slides');
    this.chapterIndicatorEl = document.getElementById('current-chapter');
    this.progressFillEl = document.getElementById('progress-fill');
    this.progressTrackEl = document.getElementById('progress-track');
    this.slideTrayModal = document.getElementById('slide-tray');
    this.trayGrid = document.getElementById('tray-grid');
    this.shortcutsDialog = document.getElementById('shortcuts-dialog');
    this.btnSound = document.getElementById('btn-sound');
    
    this.init();
  }

  init() {
    this.totalSlidesEl.textContent = String(this.totalSlides).padStart(2, '0');
    
    // Build Slide Index Grid
    this.buildSlideIndexTray();
    
    // Setup Event Listeners
    this.bindEvents();
    
    // Activate Initial Slide
    this.updateSlide(0);
  }

  bindEvents() {
    // Nav Buttons
    document.getElementById('btn-prev').addEventListener('click', () => this.prev());
    document.getElementById('btn-next').addEventListener('click', () => this.next());
    
    // HUD Buttons
    document.getElementById('btn-grid').addEventListener('click', () => this.toggleTray());
    document.getElementById('btn-fullscreen').addEventListener('click', () => this.toggleFullscreen());
    this.btnSound.addEventListener('click', () => this.toggleAmbientAudio());
    
    // Tray Close
    document.getElementById('btn-close-tray').addEventListener('click', () => this.toggleTray(false));
    document.getElementById('tray-backdrop').addEventListener('click', () => this.toggleTray(false));
    
    // Shortcuts Dialog
    document.getElementById('btn-close-dialog').addEventListener('click', () => {
      this.shortcutsDialog.classList.remove('open');
    });

    // Progress Track Seek
    this.progressTrackEl.addEventListener('click', (e) => {
      const rect = this.progressTrackEl.getBoundingClientRect();
      const clickRatio = (e.clientX - rect.left) / rect.width;
      const targetIndex = Math.floor(clickRatio * this.totalSlides);
      this.goToSlide(targetIndex);
    });

    // Keyboard Controller
    window.addEventListener('keydown', (e) => this.handleKeyDown(e));

    // Slide Tap / Click Progression
    document.getElementById('deck-viewport').addEventListener('click', (e) => {
      if (e.target.closest('video') || e.target.closest('button') || e.target.closest('.hud-btn')) return;
      this.next();
    });

    // Video Click to toggle play
    document.querySelectorAll('video').forEach(video => {
      video.addEventListener('click', (e) => {
        e.stopPropagation();
        if (video.paused) {
          video.play();
        } else {
          video.pause();
        }
      });
    });
  }

  handleKeyDown(e) {
    if (this.shortcutsDialog.classList.contains('open')) {
      if (e.key === 'Escape' || e.key === 'Enter') {
        this.shortcutsDialog.classList.remove('open');
      }
      return;
    }

    if (this.slideTrayModal.classList.contains('open')) {
      if (e.key === 'Escape' || e.key === 'Tab' || e.key.toLowerCase() === 'i') {
        this.toggleTray(false);
      }
      return;
    }

    switch (e.key) {
      case 'ArrowRight':
      case 'ArrowDown':
      case 'PageDown':
      case ' ': // Space bar
        e.preventDefault();
        this.next();
        break;

      case 'ArrowLeft':
      case 'ArrowUp':
      case 'PageUp':
        e.preventDefault();
        this.prev();
        break;

      case 'f':
      case 'F':
        e.preventDefault();
        this.toggleFullscreen();
        break;

      case 'Tab':
      case 'i':
      case 'I':
        e.preventDefault();
        this.toggleTray();
        break;

      case 'm':
      case 'M':
        e.preventDefault();
        this.toggleAmbientAudio();
        break;

      case '?':
      case '/':
        e.preventDefault();
        this.shortcutsDialog.classList.toggle('open');
        break;

      case 'Escape':
        if (document.fullscreenElement) {
          document.exitFullscreen();
        }
        break;
    }
  }

  next() {
    if (this.currentIndex < this.totalSlides - 1) {
      this.goToSlide(this.currentIndex + 1);
    }
  }

  prev() {
    if (this.currentIndex > 0) {
      this.goToSlide(this.currentIndex - 1);
    }
  }

  goToSlide(index) {
    if (index < 0 || index >= this.totalSlides || index === this.currentIndex) return;
    this.updateSlide(index);
  }

  updateSlide(newIndex) {
    const prevSlide = this.slides[this.currentIndex];
    const nextSlide = this.slides[newIndex];

    // Pause all videos on previous slide
    if (prevSlide) {
      const prevVideos = prevSlide.querySelectorAll('video');
      prevVideos.forEach(v => {
        v.pause();
        v.currentTime = 0;
      });
      prevSlide.classList.remove('active');
    }

    this.currentIndex = newIndex;

    // Activate current slide
    nextSlide.classList.add('active');

    // Autoplay videos on active slide
    const currentVideos = nextSlide.querySelectorAll('video');
    currentVideos.forEach(v => {
      v.muted = true;
      const playPromise = v.play();
      if (playPromise !== undefined) {
        playPromise.catch(err => {
          console.log('Video autoplay restrained, tap to play', err);
        });
      }
    });

    // Update Header Counter & Chapter
    this.slideNumEl.textContent = String(newIndex + 1).padStart(2, '0');
    const chapter = nextSlide.getAttribute('data-chapter') || 'RAHAT — MAKING SPACE';
    this.chapterIndicatorEl.textContent = chapter;

    // Toggle Dark Active state on app container
    const isDark = nextSlide.classList.contains('theme-dark');
    document.getElementById('presentation-app').classList.toggle('dark-active', isDark);

    // Update Progress Bar
    const progressPercent = ((newIndex + 1) / this.totalSlides) * 100;
    this.progressFillEl.style.width = `${progressPercent}%`;

    // Update Tray Active Selection
    document.querySelectorAll('.tray-thumbnail-item').forEach((item, idx) => {
      if (idx === newIndex) {
        item.classList.add('active');
      } else {
        item.classList.remove('active');
      }
    });
  }

  buildSlideIndexTray() {
    const titles = [
      "01. Rahat — Inauguration & Logo",
      "02. The Premise — Finding Space",
      "03. Evolution — Then to Now",
      "04. What We've Done — Creating Space",
      "05. What Comes Next — Expanding the Canvas",
      "06. The Scale — 700+ Packed House",
      "07. Core Committee — Leadership Team",
      "08. Connect — Community & Instagram QR",
      "09. Horizon — And This Is Only The Beginning"
    ];

    this.trayGrid.innerHTML = '';
    this.slides.forEach((slide, idx) => {
      const item = document.createElement('div');
      item.className = `tray-thumbnail-item ${idx === 0 ? 'active' : ''}`;
      item.innerHTML = `
        <div class="thumb-preview">SLIDE ${String(idx + 1).padStart(2, '0')}</div>
        <span class="thumb-num">ACT ${slide.getAttribute('data-chapter').split('—')[0].replace('ACT ', '')}</span>
        <span class="thumb-title">${titles[idx] || `Slide ${idx + 1}`}</span>
      `;
      item.addEventListener('click', () => {
        this.goToSlide(idx);
        this.toggleTray(false);
      });
      this.trayGrid.appendChild(item);
    });
  }

  toggleTray(forceState) {
    const shouldOpen = forceState !== undefined ? forceState : !this.slideTrayModal.classList.contains('open');
    if (shouldOpen) {
      this.slideTrayModal.classList.add('open');
    } else {
      this.slideTrayModal.classList.remove('open');
    }
  }

  toggleFullscreen() {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(err => {
        console.warn('Fullscreen request failed:', err);
      });
    } else {
      if (document.exitFullscreen) {
        document.exitFullscreen();
      }
    }
  }

  /**
   * Ambient Audio Engine
   */
  toggleAmbientAudio() {
    if (this.isAudioPlaying) {
      this.stopAmbientAudio();
    } else {
      this.startAmbientAudio();
    }
  }

  startAmbientAudio() {
    try {
      const AudioContext = window.AudioContext || window.webkitAudioContext;
      if (!this.audioCtx) {
        this.audioCtx = new AudioContext();
      }
      if (this.audioCtx.state === 'suspended') {
        this.audioCtx.resume();
      }

      this.audioGainNode = this.audioCtx.createGain();
      this.audioGainNode.gain.setValueAtTime(0.001, this.audioCtx.currentTime);
      this.audioGainNode.gain.exponentialRampToValueAtTime(0.12, this.audioCtx.currentTime + 3);
      this.audioGainNode.connect(this.audioCtx.destination);

      const frequencies = [110.0, 164.81, 220.0, 277.18, 329.63];
      
      this.oscillators = frequencies.map((freq, i) => {
        const osc = this.audioCtx.createOscillator();
        const panner = this.audioCtx.createStereoPanner ? this.audioCtx.createStereoPanner() : null;
        const filter = this.audioCtx.createBiquadFilter();

        osc.type = i === 0 ? 'triangle' : 'sine';
        osc.frequency.setValueAtTime(freq, this.audioCtx.currentTime);

        filter.type = 'lowpass';
        filter.frequency.setValueAtTime(450, this.audioCtx.currentTime);

        if (panner) {
          panner.pan.setValueAtTime((i % 2 === 0 ? -0.3 : 0.3), this.audioCtx.currentTime);
          osc.connect(filter);
          filter.connect(panner);
          panner.connect(this.audioGainNode);
        } else {
          osc.connect(filter);
          filter.connect(this.audioGainNode);
        }

        osc.start();
        return osc;
      });

      this.isAudioPlaying = true;
      this.btnSound.classList.add('active');
      this.btnSound.classList.remove('sound-muted');
    } catch (e) {
      console.warn('Web Audio playback error:', e);
    }
  }

  stopAmbientAudio() {
    if (this.audioGainNode && this.audioCtx) {
      this.audioGainNode.gain.exponentialRampToValueAtTime(0.0001, this.audioCtx.currentTime + 1);
      setTimeout(() => {
        this.oscillators.forEach(osc => {
          try { osc.stop(); osc.disconnect(); } catch (e) {}
        });
        this.oscillators = [];
        this.isAudioPlaying = false;
        this.btnSound.classList.remove('active');
        this.btnSound.classList.add('sound-muted');
      }, 1000);
    }
  }
}

document.addEventListener('DOMContentLoaded', () => {
  window.deck = new PresentationDeck();
});
