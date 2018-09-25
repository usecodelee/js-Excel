(function() {
    function Lcode() {}
    Lcode.prototype = {
        sumData: '',
        getData: function(e) {
            var that = this;
            var files = e.target.files;
            var fileReader = new FileReader();
            fileReader.onload = function(ev) {
                try {
                    var data = ev.target.result,
                        workbook = XLSX.read(data, {
                            type: 'binary'
                        }), // 以二进制流方式读取得到整份excel表格对象
                        persons = []; // 存储获取到的数据
                } catch (e) {
                    console.log('文件类型不正确');
                    return;
                }

                // 表格的表格范围，可用于判断表头是否数量是否正确
                var fromTo = '';
                // 遍历每张表读取
                for (var sheet in workbook.Sheets) {
                    if (workbook.Sheets.hasOwnProperty(sheet)) {
                        fromTo = workbook.Sheets[sheet]['!ref'];
                        persons = persons.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]));
                        // break; // 如果只取第一张表，就取消注释这行
                    }
                }

                that.sumData = persons;
                that.randerTable();
            };

            // 以二进制方式打开文件
            fileReader.readAsBinaryString(files[0]);
        },

        /* 求和 */
        getSum: function() {
            var that = this;
            var sum = 0;
            var data = that.sumData;
            var needSum = $('#needSum').val().trim();
            if (data == '' || needSum == '') {
                alert('文件为空或输入列名为空！');
            } else {
                for (const key in data) {
                    if (data.hasOwnProperty(key)) {
                        const element = data[key];
                        if (element[needSum]) {
                            if (isNaN(Number(element[needSum]))) {
                                alert('该列含有非数字，已经自动跳过，点击确定继续。');
                                continue;
                            }
                            sum += Number(element[needSum]);
                        } else {
                            alert('输入的列名不存在！');
                            return;
                        }

                    }
                }
                $('#result').val(sum);
            }


        },
        randerTable: function() {
            var that = this;
            var dataT = { data: that.sumData };
            var row = that.sumData.length + 1;
            $('.row1').html(row);
            var col = Object.keys(that.sumData[0]).length;
            $('.col1').html(col);
            var bt = baidu.template;
            var chtml = bt('randerT', dataT);
            $('.previewData').html(chtml);
        },

        addEvent: function() {
            var that = this;
            $('#excel-file').change(function(e) {
                that.getData(e);
            });
            $('#start').click(function(e) {
                that.getSum();
            });

        },



        init: function() {
            var that = this;
            that.addEvent();

        }
    }
    new Lcode().init();
})()