document.addEventListener('DOMContentLoaded', () => {
    // Variables
    const slides = document.querySelectorAll('.slide');
    const prevBtn = document.getElementById('prev-btn');
    const nextBtn = document.getElementById('next-btn');
    let currentSlide = 0;

    // Initialize the first slide
    slides[currentSlide].classList.add('active');

    // Function to change slide
    function changeSlide(direction) {
        // Remove active class from current slide
        slides[currentSlide].classList.remove('active');
        
        // Calculate new slide index
        if (direction === 'next') {
            currentSlide = (currentSlide + 1) % slides.length;
        } else {
            currentSlide = (currentSlide - 1 + slides.length) % slides.length;
        }
        
        // Add active class to new slide
        slides[currentSlide].classList.add('active');
    }

    // Event listeners for navigation buttons
    prevBtn.addEventListener('click', () => changeSlide('prev'));
    nextBtn.addEventListener('click', () => changeSlide('next'));

    // Keyboard navigation
    document.addEventListener('keydown', (e) => {
        if (e.key === 'ArrowRight' || e.key === ' ') {
            changeSlide('next');
        } else if (e.key === 'ArrowLeft') {
            changeSlide('prev');
        }
    });

    // Swipe navigation for touch devices
    let touchStartX = 0;
    let touchEndX = 0;
    
    document.addEventListener('touchstart', (e) => {
        touchStartX = e.changedTouches[0].screenX;
    }, false);
    
    document.addEventListener('touchend', (e) => {
        touchEndX = e.changedTouches[0].screenX;
        handleSwipe();
    }, false);
    
    function handleSwipe() {
        // Minimum distance required for swipe
        const minSwipeDistance = 50;
        
        // Calculate swipe distance
        const swipeDistance = touchEndX - touchStartX;
        
        // Check if swipe is significant enough
        if (Math.abs(swipeDistance) > minSwipeDistance) {
            if (swipeDistance > 0) {
                // Swipe right (previous)
                changeSlide('prev');
            } else {
                // Swipe left (next)
                changeSlide('next');
            }
        }
    }

    // Auto-resize text in tables based on content
    function adjustTableFontSize() {
        const tables = document.querySelectorAll('table');
        tables.forEach(table => {
            const tableWidth = table.offsetWidth;
            const contentWidth = table.scrollWidth;
            
            if (contentWidth > tableWidth) {
                const currentSize = parseInt(window.getComputedStyle(table).fontSize);
                const newSize = Math.max(currentSize * (tableWidth / contentWidth) * 0.95, 10);
                table.style.fontSize = `${newSize}px`;
            }
        });
    }

    // Call adjustTableFontSize after window loads
    window.addEventListener('load', adjustTableFontSize);
    window.addEventListener('resize', adjustTableFontSize);

    // Make charts and diagrams responsive
    function adjustImagesAndDiagrams() {
        const images = document.querySelectorAll('.architecture-diagram img, .background-image img');
        images.forEach(img => {
            // Ensure images maintain aspect ratio
            img.style.maxHeight = '100%';
            img.style.maxWidth = '100%';
        });
    }

    // Call adjustImagesAndDiagrams after window loads
    window.addEventListener('load', adjustImagesAndDiagrams);
    window.addEventListener('resize', adjustImagesAndDiagrams);
});
