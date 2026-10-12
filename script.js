/**
 * Nabihah Binti Taquddin Azmy — Portfolio Interactions
 * Accessible, Lightweight Vanilla JavaScript
 */

document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initMobileNav();
  initExperienceTabs();
  initScrollSpy();
  initScrollTop();
  initClipboardButtons();
  initNavbarScroll();
});

// 1. Theme Management (Dark / Light with persistence)
function initTheme() {
  const themeToggle = document.getElementById('themeToggle');
  const savedTheme = localStorage.getItem('nabihah_portfolio_theme');
  const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  
  // Default to obsidian dark
  const currentTheme = savedTheme || (prefersDark ? 'dark' : 'dark');
  document.documentElement.setAttribute('data-theme', currentTheme);
  updateThemeIcon(currentTheme);

  if (themeToggle) {
    themeToggle.addEventListener('click', () => {
      const activeTheme = document.documentElement.getAttribute('data-theme');
      const nextTheme = activeTheme === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', nextTheme);
      localStorage.setItem('nabihah_portfolio_theme', nextTheme);
      updateThemeIcon(nextTheme);
    });
  }
}

function updateThemeIcon(theme) {
  const themeToggle = document.getElementById('themeToggle');
  if (!themeToggle) return;
  
  if (theme === 'dark') {
    // Sun icon to toggle light mode
    themeToggle.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="5"></circle>
        <line x1="12" y1="1" x2="12" y2="3"></line>
        <line x1="12" y1="21" x2="12" y2="23"></line>
        <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line>
        <line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line>
        <line x1="1" y1="12" x2="3" y2="12"></line>
        <line x1="21" y1="12" x2="23" y2="12"></line>
        <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line>
        <line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line>
      </svg>
    `;
    themeToggle.setAttribute('aria-label', 'Switch to light mode');
  } else {
    // Moon icon to toggle dark mode
    themeToggle.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path>
      </svg>
    `;
    themeToggle.setAttribute('aria-label', 'Switch to dark mode');
  }
}

// 2. Mobile Navigation Drawer
function initMobileNav() {
  const btnOpen = document.getElementById('btnMobileNav');
  const btnClose = document.getElementById('btnCloseMobileNav');
  const drawer = document.getElementById('mobileNavDrawer');
  const overlay = document.getElementById('mobileNavOverlay');
  const mobileLinks = document.querySelectorAll('.mobile-nav-link');

  const openDrawer = () => {
    if (drawer) {
      drawer.classList.add('active');
      drawer.setAttribute('aria-hidden', 'false');
    }
    if (overlay) {
      overlay.classList.add('active');
      overlay.setAttribute('aria-hidden', 'false');
    }
    document.body.style.overflow = 'hidden';
  };

  const closeDrawer = () => {
    if (drawer) {
      drawer.classList.remove('active');
      drawer.setAttribute('aria-hidden', 'true');
    }
    if (overlay) {
      overlay.classList.remove('active');
      overlay.setAttribute('aria-hidden', 'true');
    }
    document.body.style.overflow = '';
  };

  if (btnOpen) btnOpen.addEventListener('click', openDrawer);
  if (btnClose) btnClose.addEventListener('click', closeDrawer);
  if (overlay) overlay.addEventListener('click', closeDrawer);

  mobileLinks.forEach(link => {
    link.addEventListener('click', closeDrawer);
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && drawer && drawer.classList.contains('active')) {
      closeDrawer();
    }
  });
}

// 3. Experience Interactive Tabs & Mobile Dropdown
function initExperienceTabs() {
  const tabBtns = document.querySelectorAll('.tab-btn');
  const panels = document.querySelectorAll('.tab-panel');
  const expSelect = document.getElementById('expSelect');

  if (!panels.length) return;

  const switchTab = (targetId) => {
    // Update button active state & ARIA
    tabBtns.forEach(btn => {
      const isSelected = btn.getAttribute('data-tab') === targetId;
      btn.classList.toggle('active', isSelected);
      btn.setAttribute('aria-selected', isSelected ? 'true' : 'false');
    });

    // Sync mobile select element
    if (expSelect && expSelect.value !== targetId) {
      expSelect.value = targetId;
    }

    // Update panel visibility
    panels.forEach(panel => {
      panel.classList.toggle('active', panel.id === targetId);
    });
  };

  // Click on desktop tab button
  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const targetId = btn.getAttribute('data-tab');
      if (targetId) switchTab(targetId);
    });
  });

  // Change on mobile select dropdown
  if (expSelect) {
    expSelect.addEventListener('change', (e) => {
      const targetId = e.target.value;
      if (targetId) switchTab(targetId);
    });
  }
}

// 4. Scroll Spy (Active nav indicator)
function initScrollSpy() {
  const sections = document.querySelectorAll('section[id]');
  const navLinks = document.querySelectorAll('.nav-link');
  const mobileNavLinks = document.querySelectorAll('.mobile-nav-link');

  window.addEventListener('scroll', () => {
    let currentId = '';
    const scrollPos = window.scrollY + 160;

    sections.forEach(sec => {
      const top = sec.offsetTop;
      const height = sec.offsetHeight;
      if (scrollPos >= top && scrollPos < top + height) {
        currentId = sec.getAttribute('id');
      }
    });

    navLinks.forEach(link => {
      link.classList.remove('active');
      if (link.getAttribute('href') === `#${currentId}`) {
        link.classList.add('active');
      }
    });

    mobileNavLinks.forEach(link => {
      link.classList.remove('active');
      if (link.getAttribute('href') === `#${currentId}`) {
        link.classList.add('active');
      }
    });
  }, { passive: true });
}

// 5. Scroll to Top Button
function initScrollTop() {
  const btnScrollTop = document.getElementById('btnScrollTop');
  if (!btnScrollTop) return;

  window.addEventListener('scroll', () => {
    if (window.scrollY > 400) {
      btnScrollTop.classList.add('visible');
    } else {
      btnScrollTop.classList.remove('visible');
    }
  }, { passive: true });

  btnScrollTop.addEventListener('click', () => {
    window.scrollTo({
      top: 0,
      behavior: 'smooth'
    });
  });
}

// 6. Clipboard Copy Feedback
function initClipboardButtons() {
  const copyButtons = document.querySelectorAll('[data-copy]');
  copyButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      const text = btn.getAttribute('data-copy');
      if (!text) return;

      navigator.clipboard.writeText(text).then(() => {
        showToast(`Copied "${text}" to clipboard!`);
      }).catch(() => {
        showToast('Failed to copy to clipboard');
      });
    });
  });
}

function showToast(message) {
  let toast = document.querySelector('.toast-bar');
  if (!toast) {
    toast = document.createElement('div');
    toast.className = 'toast-bar';
    document.body.appendChild(toast);
  }

  toast.textContent = message;
  toast.classList.add('active');

  setTimeout(() => {
    toast.classList.remove('active');
  }, 2600);
}

// 7. Navbar elevation on scroll
function initNavbarScroll() {
  const navbar = document.getElementById('navbar');
  if (!navbar) return;

  window.addEventListener('scroll', () => {
    if (window.scrollY > 30) {
      navbar.classList.add('scrolled');
    } else {
      navbar.classList.remove('scrolled');
    }
  }, { passive: true });
}
