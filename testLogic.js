const item = { "bud ADV MONTOK": 5, "actAdvMontok": 10 };
const hybridList = ["ADV MONTOK", "ADV JOSS", "ADV BEJO", "ADV GANESH", "ADV JAGO", "ADV JALU", "ADV RUBY"];
hybridList.forEach(cropName => {
    let budProp = "";
    let actProp = "";
    Object.keys(item).forEach(k => {
        const cleanK = k.toLowerCase().replace(/[^a-z0-9]/g, '');
        const cleanCrop = cropName.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (cleanK === `bud${cleanCrop}`) budProp = k;
        if (cleanK === `act${cleanCrop}`) actProp = k;
    });
    let budVal = budProp ? Number(item[budProp] || 0) : 0;
    let actVal = actProp ? Number(item[actProp] || 0) : 0;
    if (budProp || actProp) {
        console.log(`Found ${cropName}: bud=${budVal}, act=${actVal}`);
    }
});
