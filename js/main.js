/* ============================================
   DESTILERIJA RANKOVIĆ - Main JavaScript
   ============================================ */

document.addEventListener('DOMContentLoaded', function()
{
    /* ============================================
       AGE VERIFICATION
       ============================================ */
    const ageModal = document.getElementById('ageModal');
    const ageYes = document.getElementById('ageYes');
    const ageNo = document.getElementById('ageNo');

    // Check if user already verified
    const isVerified = localStorage.getItem('ageVerified');

    if (isVerified === 'true')
    {
        ageModal.classList.add('hidden');
        document.body.style.overflow = '';
    }
    else
    {
        document.body.style.overflow = 'hidden';
    }

    // Yes button - allow access
    ageYes.addEventListener('click', function()
    {
        localStorage.setItem('ageVerified', 'true');
        ageModal.classList.add('hidden');
        document.body.style.overflow = '';
    });

    // No button - redirect away
    ageNo.addEventListener('click', function()
    {
        window.location.href = 'https://www.google.com';
    });

    // Elements
    const header = document.getElementById('header');
    const nav = document.getElementById('nav');
    const hamburger = document.getElementById('hamburger');
    const navLinks = document.querySelectorAll('.nav-link');
    const galleryItems = document.querySelectorAll('.gallery-item');
    const lightbox = document.getElementById('lightbox');
    const lightboxImage = document.getElementById('lightboxImage');
    const lightboxClose = document.getElementById('lightboxClose');
    const contactForm = document.getElementById('contactForm');
    const revealElements = document.querySelectorAll('.reveal');

    /* ============================================
       HEADER SCROLL EFFECT
       ============================================ */
    function handleScroll()
    {
        if (window.scrollY > 100)
        {
            header.classList.add('scrolled');
        }
        else
        {
            header.classList.remove('scrolled');
        }
    }

    window.addEventListener('scroll', handleScroll);
    handleScroll(); // Check on load

    /* ============================================
       MOBILE MENU
       ============================================ */
    hamburger.addEventListener('click', function()
    {
        hamburger.classList.toggle('active');
        nav.classList.toggle('active');
        document.body.style.overflow = nav.classList.contains('active') ? 'hidden' : '';
    });

    // Close menu on link click
    navLinks.forEach(function(link)
    {
        link.addEventListener('click', function()
        {
            hamburger.classList.remove('active');
            nav.classList.remove('active');
            document.body.style.overflow = '';
        });
    });

    /* ============================================
       SMOOTH SCROLL
       ============================================ */
    document.querySelectorAll('a[href^="#"]').forEach(function(anchor)
    {
        anchor.addEventListener('click', function(e)
        {
            const href = this.getAttribute('href');

            if (href !== '#')
            {
                e.preventDefault();
                const target = document.querySelector(href);

                if (target)
                {
                    const headerHeight = header.offsetHeight;
                    const targetPosition = target.offsetTop - headerHeight;

                    window.scrollTo({
                        top: targetPosition,
                        behavior: 'smooth'
                    });
                }
            }
        });
    });

    /* ============================================
       SCROLL REVEAL ANIMATIONS
       ============================================ */
    const revealObserver = new IntersectionObserver(function(entries)
    {
        entries.forEach(function(entry)
        {
            if (entry.isIntersecting)
            {
                entry.target.classList.add('active');
            }
        });
    }, {
        threshold: 0.1,
        rootMargin: '0px 0px -50px 0px'
    });

    revealElements.forEach(function(element)
    {
        revealObserver.observe(element);
    });

    /* ============================================
       GALLERY LIGHTBOX
       ============================================ */
    galleryItems.forEach(function(item)
    {
        item.addEventListener('click', function()
        {
            const img = this.querySelector('img');
            lightboxImage.src = img.src;
            lightboxImage.alt = img.alt;
            lightbox.classList.add('active');
            document.body.style.overflow = 'hidden';
        });
    });

    function closeLightbox()
    {
        lightbox.classList.remove('active');
        document.body.style.overflow = '';
    }

    lightboxClose.addEventListener('click', closeLightbox);

    lightbox.addEventListener('click', function(e)
    {
        if (e.target === lightbox)
        {
            closeLightbox();
        }
    });

    document.addEventListener('keydown', function(e)
    {
        if (e.key === 'Escape' && lightbox.classList.contains('active'))
        {
            closeLightbox();
        }
    });

    /* ============================================
       CONTACT FORM
       ============================================ */
    contactForm.addEventListener('submit', function(e)
    {
        e.preventDefault();

        const formData = new FormData(this);
        const data = {};

        formData.forEach(function(value, key)
        {
            data[key] = value;
        });

        // Show success message
        const button = this.querySelector('.btn-submit');
        const originalText = button.textContent;

        button.textContent = 'Poruka Poslata!';
        button.style.background = '#28a745';
        button.style.borderColor = '#28a745';

        // Reset form
        this.reset();

        // Restore button after 3 seconds
        setTimeout(function()
        {
            button.textContent = originalText;
            button.style.background = '';
            button.style.borderColor = '';
        }, 3000);
    });

    /* ============================================
       ACTIVE NAV LINK ON SCROLL
       ============================================ */
    const sections = document.querySelectorAll('section[id]');

    function setActiveNavLink()
    {
        const scrollY = window.scrollY;

        sections.forEach(function(section)
        {
            const sectionHeight = section.offsetHeight;
            const sectionTop = section.offsetTop - 150;
            const sectionId = section.getAttribute('id');

            if (scrollY > sectionTop && scrollY <= sectionTop + sectionHeight)
            {
                navLinks.forEach(function(link)
                {
                    link.classList.remove('active');

                    if (link.getAttribute('href') === '#' + sectionId)
                    {
                        link.classList.add('active');
                    }
                });
            }
        });
    }

    window.addEventListener('scroll', setActiveNavLink);

    /* ============================================
       PARALLAX EFFECT FOR HERO (subtle)
       ============================================ */
    const heroBg = document.querySelector('.hero-bg');

    if (heroBg)
    {
        window.addEventListener('scroll', function()
        {
            const scrolled = window.scrollY;

            if (scrolled < window.innerHeight)
            {
                heroBg.style.transform = 'translateY(' + (scrolled * 0.3) + 'px)';
            }
        });
    }
});
