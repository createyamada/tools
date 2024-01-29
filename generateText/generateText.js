window.onload = function() {
    // ページ読み込み時に実行したい処理
    let array = [];
    for( let i=0;i<10;i++ ) {
        // 偶数の時だけ
        if( i % 2 === 0 ) {
            let test = "入れたい文字列を挿入";
        }
        array.push(`
            繰り返し文字列を入力${i}
        `);
    }

    let blob = new Blob(array,{type:"text/plan"});
    let link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = '作ったファイル.txt';
    link.click();
}