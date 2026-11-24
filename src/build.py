t=open('/home/claude/nt/tpl2.html',encoding='utf8').read()
off=open('/mnt/user-data/outputs/Ton_kho_cac_chi_nhanh.html',encoding='utf8').read()
i=off.index('/*! xlsx.js');lib=off[i:off.index('</script>',i)]
core=open('/home/claude/nt/core.js',encoding='utf8').read().replace("if(typeof module!=='undefined')module.exports=NT;","")
def b2(libhtml,out):
    x=t.split('<script>/*LIB*/</script>');x=x[0]+libhtml+x[1];u=x.split('/*CORE*/');open(out,'w',encoding='utf8').write(u[0]+core+u[1])
b2('<script>'+lib+'</script>','/mnt/user-data/outputs/Ton_kho_cac_chi_nhanh.html')
b2('<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>','/mnt/user-data/outputs/AppsScript/Index.html')
print('built')
