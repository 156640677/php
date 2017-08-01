/**   
 * json��ʽת��״�ṹ   
 * @param   {json}      json����   
 * @param   {String}    id���ַ���   
 * @param   {String}    ��id���ַ���   
 * @param   {String}    children���ַ���   
 * @return  {Array}     ����   
 */    
function transData(a, idStr, pidStr, chindrenStr){    
    var r = [], hash = {}, id = idStr, pid = pidStr, children = chindrenStr, i = 0, j = 0, len = a.length;    
    for(; i < len; i++){    
        hash[a[i][id]] = a[i];    
    }    
    for(; j < len; j++){    
        var aVal = a[j], hashVP = hash[aVal[pid]];    
        if(hashVP){    
            !hashVP[children] && (hashVP[children] = []);    
            hashVP[children].push(aVal);    
        }else{    
            r.push(aVal);    
        }    
    }    
    return r;    
}    
    
var jsonData = eval('[  
    {"id":"4","pid":"1","name":"��ҵ�"},  
    {"id":"5","pid":"1","name":"�������"},  
    {"id":"1","pid":"0","name":"���õ���"},  
    {"id":"2","pid":"0","name":"����"},  
    {"id":"3","pid":"0","name":"��ױ"},  
    {"id":"7","pid":"4","name":"�յ�"},  
    {"id":"8","pid":"4","name":"����"},  
    {"id":"9","pid":"4","name":"ϴ�»�"},  
    {"id":"10","pid":"4","name":"��ˮ��"},  
    {"id":"11","pid":"3","name":"�沿����"},  
    {"id":"12","pid":"3","name":"��ǻ����"},  
    {"id":"13","pid":"2","name":"��װ"},  
    {"id":"14","pid":"2","name":"Ůװ"},  
    {"id":"15","pid":"7","name":"�����յ�"},  
    {"id":"16","pid":"7","name":"���Ŀյ�"},  
    {"id":"19","pid":"5","name":"��ʪ��"},  
    {"id":"20","pid":"5","name":"���ٶ�"}  
    ]');    
    
var jsonDataTree = transData(jsonData, 'id', 'pid', 'chindren');    
console.log(jsonDataTree);    
//������£�  
[  
    {"id":"1","pid":"0","name":"���õ���", "chindren":[  
        {"id":"4","pid":"1","name":"��ҵ�", "chindren":[  
            {"id":"7","pid":"4","name":"�յ�", "chindren":[  
                {"id":"15","pid":"7","name":"�����յ�"},  
                {"id":"16","pid":"7","name":"���Ŀյ�"}  
            ]},  
            {"id":"8","pid":"4","name":"����"},  
            {"id":"9","pid":"4","name":"ϴ�»�"},  
            {"id":"10","pid":"4","name":"��ˮ��"}  
        ]},  
        {"id":"5","pid":"1","name":"�������","chindren":[  
            {"id":"19","pid":"5","name":"��ʪ��"},  
            {"id":"20","pid":"5","name":"���ٶ�"}  
        ]}  
    ]},  
    {"id":"2","pid":"0","name":"����","chindren":[  
        {"id":"13","pid":"2","name":"��װ"},  
        {"id":"14","pid":"2","name":"Ůװ"}  
    ]},  
    {"id":"3","pid":"0","name":"��ױ","chindren":[  
        {"id":"11","pid":"3","name":"�沿����"},  
        {"id":"12","pid":"3","name":"��ǻ����"}  
    ]}  
]    