interface Divs2SlidesOptions {
  first?: number;
  nav?: boolean;
  showPlayPauseBtn?: boolean;
  showFullscreenBtn?: boolean;
  navTxtColor?: string;
  keyBoardShortCut?: boolean;
  showSlideNum?: boolean;
  showTotalSlideNum?: boolean;
  autoSlide?: number | false;
  randomAutoSlide?: boolean;
  loop?: boolean;
  background?: string | string[] | false;
  transition?: 'default' | 'fade' | 'slide' | 'random';
  transitionTime?: number;
}

class Divs2Slides {
  private options: Divs2SlidesOptions;
  private container: HTMLElement;
  private slides: HTMLElement[];
  private currentSlide: number;
  private totalSlides: number;
  private isSlideMode: boolean;
  private isFullscreen: boolean;
  private autoSlideInterval: number | null;

  constructor(container: HTMLElement, options: Divs2SlidesOptions = {}) {
    this.options = {
      first: 1,
      nav: true,
      showPlayPauseBtn: true,
      showFullscreenBtn: true,
      navTxtColor: 'black',
      keyBoardShortCut: true,
      showSlideNum: true,
      showTotalSlideNum: true,
      autoSlide: false,
      randomAutoSlide: false,
      loop: false,
      background: false,
      transition: 'default',
      transitionTime: 1,
      ...options
    };

    this.container = container;
    this.slides = Array.from(this.container.querySelectorAll('.slide'));
    this.currentSlide = this.options.first ? this.options.first - 1 : 0;
    this.totalSlides = this.slides.length;
    this.isSlideMode = true;
    this.isFullscreen = false;
    this.autoSlideInterval = null;

    this.init();

    document.addEventListener('fullscreenchange', this.handleFullscreenChange.bind(this));
  }

  private init(): void {
    this.setupSlides();
    this.createNavigation();
    if (this.options.keyBoardShortCut) {
      document.addEventListener('keydown', this.handleKeyDown.bind(this));
    }
    this.goToSlide(this.currentSlide);
  }

  private setupSlides(): void {
    this.slides.forEach((slide, index) => {
      slide.style.display = 'none';
      if (this.options.background) {
        if (Array.isArray(this.options.background)) {
          slide.style.backgroundColor = this.options.background[index % this.options.background.length];
        } else if (typeof this.options.background === 'string') {
          slide.style.backgroundColor = this.options.background;
        }
      }
    });
  }

  private createNavigation(): void {
    if (!this.options.nav) return;

    const toolbar = document.createElement('div');
    toolbar.className = 'slides-toolbar';
    toolbar.style.cssText = 'width: 90%; padding: 10px; text-align: center; position: fixed; bottom: 0; left: 5%;';

    if (this.options.showPlayPauseBtn) {
      const playPauseBtn = this.createButton('▶', this.toggleAutoSlide.bind(this));
      toolbar.appendChild(playPauseBtn);
    }

    if (this.options.showFullscreenBtn) {
      const fullscreenBtn = this.createButton('⤢', this.toggleFullscreen.bind(this));
      toolbar.appendChild(fullscreenBtn);
    }

    const prevBtn = this.createButton('←', this.prevSlide.bind(this));
    const nextBtn = this.createButton('→', this.nextSlide.bind(this));

    toolbar.appendChild(prevBtn);
    toolbar.appendChild(nextBtn);

    if (this.options.showSlideNum || this.options.showTotalSlideNum) {
      const slideInfo = document.createElement('span');
      slideInfo.style.marginLeft = '10px';
      slideInfo.textContent = `${this.options.showSlideNum ? this.currentSlide + 1 : ''}${this.options.showTotalSlideNum ? ' / ' + this.totalSlides : ''}`;
      toolbar.appendChild(slideInfo);
    }

    this.container.appendChild(toolbar);
  }

  private createButton(text: string, onClick: () => void): HTMLButtonElement {
    const button = document.createElement('button');
    button.textContent = text;
    button.style.cssText = 'margin: 0 5px; padding: 5px 10px;';
    button.addEventListener('click', onClick);
    return button;
  }

  private handleKeyDown(event: KeyboardEvent): void {
    switch (event.key) {
      case 'ArrowLeft':
        this.prevSlide();
        break;
      case 'ArrowRight':
      case ' ':
        this.nextSlide();
        break;
      case 'f':
        this.toggleFullscreen();
        break;
    }
  }

  private goToSlide(index: number): void {
    if (index < 0 || index >= this.totalSlides) return;

    const fromSlide = this.slides[this.currentSlide];
    const toSlide = this.slides[index];

    this.applyTransition(fromSlide, toSlide);
    this.currentSlide = index;
    this.updateNavigation();
  }

  private updateNavigation(): void {
    if (!this.options.nav) return;

    const slideInfo = this.container.querySelector('.slides-toolbar span');
    if (slideInfo) {
      slideInfo.textContent = `${this.options.showSlideNum ? this.currentSlide + 1 : ''}${this.options.showTotalSlideNum ? ' / ' + this.totalSlides : ''}`;
    }
  }

  public nextSlide(): void {
    let nextIndex = this.currentSlide + 1;
    if (nextIndex >= this.totalSlides) {
      nextIndex = this.options.loop ? 0 : this.currentSlide;
    }
    this.goToSlide(nextIndex);
  }

  public prevSlide(): void {
    let prevIndex = this.currentSlide - 1;
    if (prevIndex < 0) {
      prevIndex = this.options.loop ? this.totalSlides - 1 : 0;
    }
    this.goToSlide(prevIndex);
  }

  private toggleAutoSlide(): void {
    if (this.autoSlideInterval) {
      clearInterval(this.autoSlideInterval);
      this.autoSlideInterval = null;
    } else if (typeof this.options.autoSlide === 'number') {
      this.autoSlideInterval = window.setInterval(() => {
        if (this.currentSlide === this.totalSlides - 1 && !this.options.loop) {
          this.stopAutoSlide();
        } else {
          this.nextSlide();
        }
      }, this.options.autoSlide * 1000);
    }
    this.updatePlayPauseButton();
  }

  private stopAutoSlide(): void {
    if (this.autoSlideInterval) {
      clearInterval(this.autoSlideInterval);
      this.autoSlideInterval = null;
    }
    this.updatePlayPauseButton();
  }

  private updatePlayPauseButton(): void {
    const playPauseBtn = this.container.querySelector('.play-pause-btn') as HTMLButtonElement;
    if (playPauseBtn) {
      playPauseBtn.textContent = this.autoSlideInterval ? '⏸' : '▶';
    }
  }

  private toggleFullscreen(): void {
    if (!this.isFullscreen) {
      if (this.container.requestFullscreen) {
        this.container.requestFullscreen();
      } else if ((this.container as any).mozRequestFullScreen) {
        (this.container as any).mozRequestFullScreen();
      } else if ((this.container as any).webkitRequestFullscreen) {
        (this.container as any).webkitRequestFullscreen();
      } else if ((this.container as any).msRequestFullscreen) {
        (this.container as any).msRequestFullscreen();
      }
    } else {
      if (document.exitFullscreen) {
        document.exitFullscreen();
      } else if ((document as any).mozCancelFullScreen) {
        (document as any).mozCancelFullScreen();
      } else if ((document as any).webkitExitFullscreen) {
        (document as any).webkitExitFullscreen();
      } else if ((document as any).msExitFullscreen) {
        (document as any).msExitFullscreen();
      }
    }
    this.isFullscreen = !this.isFullscreen;
  }

  private applyTransition(fromSlide: HTMLElement, toSlide: HTMLElement): void {
    const transitionTime = this.options.transitionTime || 1;
    let transition = this.options.transition || 'default';
    
    if (transition === 'random') {
      const transitions: Array<'default' | 'fade' | 'slide'> = ['default', 'fade', 'slide'];
      transition = transitions[Math.floor(Math.random() * transitions.length)];
    }

    fromSlide.style.transition = `all ${transitionTime}s`;
    toSlide.style.transition = `all ${transitionTime}s`;

    switch (transition) {
      case 'fade':
        fromSlide.style.opacity = '0';
        toSlide.style.opacity = '1';
        break;
      case 'slide':
        fromSlide.style.transform = 'translateX(-100%)';
        toSlide.style.transform = 'translateX(0)';
        break;
      default:
        fromSlide.style.display = 'none';
        toSlide.style.display = 'block';
    }

    setTimeout(() => {
      fromSlide.style.display = 'none';
      toSlide.style.display = 'block';
      fromSlide.style.opacity = '1';
      fromSlide.style.transform = 'translateX(0)';
    }, transitionTime * 1000);
  }

  private handleFullscreenChange(): void {
    this.isFullscreen = !!document.fullscreenElement;
    if (this.isFullscreen) {
      this.container.classList.add('fullscreen');
      this.resizeSlides();
    } else {
      this.container.classList.remove('fullscreen');
      this.resetSlidesSize();
    }
  }

  private resizeSlides(): void {
    const containerWidth = this.container.clientWidth;
    const containerHeight = this.container.clientHeight;
    const scale = Math.min(containerWidth / 1024, containerHeight / 768); // Assuming 1024x768 base size
    this.slides.forEach(slide => {
      slide.style.transform = `scale(${scale})`;
      slide.style.transformOrigin = 'top left';
    });
  }

  private resetSlidesSize(): void {
    this.slides.forEach(slide => {
      slide.style.transform = '';
      slide.style.transformOrigin = '';
    });
  }

  public toggleSlideMode(): void {
    this.isSlideMode = !this.isSlideMode;
    if (this.isSlideMode) {
      this.container.classList.add('slide-mode');
      this.goToSlide(this.currentSlide);
    } else {
      this.container.classList.remove('slide-mode');
      this.slides.forEach(slide => slide.style.display = 'block');
      this.stopAutoSlide();
    }
    this.updateNavigation();
  }
}

// original slidemode css
// if (this.options.slideMode && this.options.slideType == "divs2slidesjs") {
//     cssText += "#all_slides_warpper{margin-right: auto;margin-left: auto;padding-top:10px;width: " + this.basicInfo?.width + "px;}\n";
// }

// Usage:
// const container = document.getElementById('slideContainer');
// const options = { /* your options here */ };
// const slideshow = new Divs2Slides(container, options);

export default Divs2Slides;
