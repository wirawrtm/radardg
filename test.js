const url = "https://script.google.com/macros/s/AKfycbxVCPzllBhLd8J-1UWeDJteFshrCs2M2PtehcXj7mHpW3PWfAcXe1d69NSkE9j3LYnM7A/exec?action=getOverviewData";
fetch(url).then(r => r.json()).then(d => console.log(d.data.length)).catch(e => console.error(e));
