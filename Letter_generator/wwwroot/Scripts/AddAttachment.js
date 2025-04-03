document.addEventListener("DOMContentLoaded", function () {
    let appCounter = 0;
    const container = document.getElementById("applications-container");

    function createApplicationBlock(index) {
        const div = document.createElement("div");
        div.classList.add("application-block");
        div.innerHTML = `
            <h3>Приложение ${index + 1}</h3>
            <input type="text" name="Attachments[${index}].Theme" placeholder="Тема приложения" >
            <textarea name="Attachments[${index}].Text" placeholder="Текст приложения" ></textarea>
        `;
        container.appendChild(div);

        const inputs = div.querySelectorAll("input, textarea");
        inputs.forEach(input => {
            input.addEventListener("input", checkLastApplication);
        });
    }

    function checkLastApplication() {
        const lastAppInputs = container.lastElementChild.querySelectorAll("input, textarea");
        const allFilled = Array.from(lastAppInputs).every(input => input.value.trim() !== "");
        if (allFilled) {
            appCounter++;
            createApplicationBlock(appCounter);
        }
    }

    createApplicationBlock(appCounter);
});