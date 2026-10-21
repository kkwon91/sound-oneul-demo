const template = document.getElementById('inquiry-template');
const copyButton = document.getElementById('copy-inquiry');
const copyStatus = document.getElementById('inquiry-status');

if (template && copyButton && copyStatus) {
  copyButton.hidden = false;
  copyButton.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(template.value);
      copyStatus.textContent = '복사했습니다. 카카오톡이나 인스타그램 DM에 붙여 넣고 작성해 주세요.';
    } catch {
      template.focus();
      template.select();
      template.setSelectionRange(0, template.value.length);
      copyStatus.textContent = '양식을 선택했습니다. 직접 복사해서 카카오톡이나 인스타그램 DM에 붙여 넣어주세요.';
    }
  });
}
