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

    // Form submission handling via fetch
    const form = document.getElementById('contactForm');
    if (form) {
        form.addEventListener('submit', (e) => {
            e.preventDefault();
            const btn = form.querySelector('.btn-primary');
            const originalText = btn.textContent;
            
            // Visual feedback
            btn.textContent = 'Enviando...';
            btn.style.opacity = '0.7';
            
            const formData = new FormData(form);
            
            fetch(form.action, {
                method: "POST",
                body: formData,
                headers: {
                    'Accept': 'application/json'
                }
            }).then(response => {
                if (response.ok) {
                    btn.textContent = '¡Mensaje Enviado!';
                    btn.style.backgroundColor = '#22c55e'; // Success green
                    btn.style.opacity = '1';
                    form.reset();
                    
                    setTimeout(() => {
                        btn.textContent = originalText;
                        btn.style.backgroundColor = '';
                    }, 3000);
                } else {
                    btn.textContent = 'Error al enviar';
                    btn.style.backgroundColor = '#ef4444'; // Error red
                    
                    setTimeout(() => {
                        btn.textContent = originalText;
                        btn.style.backgroundColor = '';
                        btn.style.opacity = '1';
                    }, 3000);
                }
            }).catch(error => {
                console.error('Error:', error);
                btn.textContent = 'Error al enviar';
                btn.style.backgroundColor = '#ef4444'; // Error red
            });
        });
    }
});
