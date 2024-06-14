// // Generate random color for gradient
// function getRandomColor() {
//     const letters = '0123456789ABCDEF';
//     let color = '#';
//     for (let i = 0; i < 6; i++) {
//         color += letters[Math.floor(Math.random() * 16)];
//     }
//     return color;
// }

// // Generate random gradient
// function getRandomGradient() {
//     const color1 = getRandomColor();
//     const color2 = getRandomColor();
//     return `linear-gradient(160deg, ${color1} 0%, ${color2} 100%)`;
// }

// // Change background gradient with fade
// function changeBackgroundGradient() {
//     const newGradient = getRandomGradient();

//     const overlay = document.createElement('div');
//     overlay.style.position = 'fixed';
//     overlay.style.top = '0';
//     overlay.style.left = '0';
//     overlay.style.width = '100vw';
//     overlay.style.height = '100vh';
//     overlay.style.backgroundImage = newGradient;
//     overlay.style.opacity = '0';
//     overlay.style.transition = 'opacity 2s ease';

//     document.body.appendChild(overlay);

//     overlay.offsetHeight;

//     overlay.style.opacity = '1';

//     setTimeout(() => {
//         document.body.style.backgroundImage = newGradient;
//         overlay.remove();
//     }, 2000); 
// }

// // Change the background gradient every 5 seconds
// setInterval(changeBackgroundGradient, 8000);
