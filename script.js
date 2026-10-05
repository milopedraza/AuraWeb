document.addEventListener('DOMContentLoaded', () => {
    // Set current year in footer
    document.getElementById('year').textContent = new Date().getFullYear();

    // Smooth scrolling for navigation links
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const targetId = this.getAttribute('href');
            if (targetId === '#') return;
            
            const targetElement = document.querySelector(targetId);
            if (targetElement) {
                targetElement.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });

    // Form submission handling (Mockup)
    const form = document.getElementById('contactForm');
    if (form) {
        form.addEventListener('submit', (e) => {
            e.preventDefault();
            const btn = form.querySelector('.btn-submit');
            const originalText = btn.textContent;
            
            // Visual feedback
            btn.textContent = 'Enviando...';
            btn.style.opacity = '0.7';
            
            // Simulate API call
            setTimeout(() => {
                btn.textContent = '¡Mensaje Enviado!';
                btn.style.backgroundColor = '#22c55e'; // Success green
                btn.style.opacity = '1';
                form.reset();
                
                // Reset button after 3 seconds
                setTimeout(() => {
                    btn.textContent = originalText;
                    btn.style.backgroundColor = '';
                }, 3000);
            }, 1500);
        });
    }
});
