const Vue = require('./node_modules/vue/dist/vue.js');
const ElementUI = require('element-ui');
const rasterizeHTML = require('rasterizehtml');
const XLSX = require('xlsx');
const { dialog } = require('electron').remote;
const { shell } = require('electron');
const fs = require('fs');
const chokidar = require('chokidar');

Vue.use(ElementUI);
const { ipcRenderer } = require('electron');
const htmlStyles = `<style> body{-webkit-touch-callout:none;font-family:-apple-system-font,BlinkMacSystemFont,"Helvetica Neue","PingFang SC","Hiragino Sans GB","Microsoft YaHei UI","Microsoft YaHei",Arial,sans-serif;color:#000;letter-spacing:.034em; padding:0; margin:0;} </style>`;
new Vue({
  el: '#app',
  data: function() {
    return {
      visible: false,
      filePath: '',
      html: '',
      fields: [],
      fieldsPosition: [],
      tableData: [],
      cellTypes: {},
      currentPage: 'import',
      exportType: '',
      exportDir: '',
      stampDir: '',
      loading: false,
      progress: {
        total: 0,
        finished: 0
      },
      cfg: {
        types: []
      },
      watcher: null
    }
  },
  watch: {
    stampDir() {
      this.readDocCfg();
    }
  },
  computed: {
    hasReady() {
      return (!!this.exportType && !!this.filePath && !!this.exportDir && !!this.stampDir);
    },
    done() {
      return this.progress.total > 0 && 
        this.progress.finished === this.progress.total;
    },
    selectedType() {
      if(!this.cfg.types) {
        return;
      }
      for(let type of this.cfg.types) {
        if (type.name === this.exportType) {
          return type;
        }
      }
      return {};
    },
    types() {
      const types = [];
      for(let type of this.cfg.types) {
        types.push({
          label: type.label,
          value: type.name
        });
      }
      return types;
    }
  },
  methods: {
    getAppPath() {
      return new Promise((resolve, reject) => {
        ipcRenderer.send('get-app-path');
        ipcRenderer.on('got-app-path', (event, path) => {
          resolve(path)
        });
      });
    },
    getTypeCfg() {
      if(!this.cfg.types) {
        return;
      }
      for(let type of this.cfg.types) {
        if (type.name === this.exportType) {
          return type;
        }
      }
      return {};
    },
    getPicFile(dir, name) {
      const file = `${dir}/${name}.png`;
      if(fs.existsSync(file)) {
        return file;
      } else {
        const files = fs.readdirSync(dir);
        const filtedFiles = files.filter(file => file.startsWith(name));
        const randomIdx = Math.floor(Math.random() * filtedFiles.length);
        return `${dir}/${filtedFiles[randomIdx]}`;
      }
    },
    toImg() {
      var canvas = document.getElementById("canvas");
      canvas.width = 1000;
      canvas.height = 3000;
      rasterizeHTML.drawHTML(this.html, canvas);
    },
    async convert() {
      this.readDocCfg();
      this.progress.total = 0;
      this.progress.finished = 0;
      //this.cfg.types[1].pages[0].placeholders[0] = ;
      this.loading = true;
      const stampDir = this.stampDir;
      const exportDir = this.exportDir;
      const typeCfg = this.selectedType;
      console.log('total', this.tableData.length * typeCfg.pages.length);
      this.progress.total = this.tableData.length * typeCfg.pages.length;
      const dirPrefix = this.getDirPrefix();
      for(let pageCfg of typeCfg.pages) {
        for(let i in this.tableData) {
          let item = this.tableData[i];
          const res = await this.convertPage({
            i,
            dir: stampDir,
            pageCfg,
            row: item
          });
          const destDir = `/${dirPrefix}/` + this.createName(item);
          this.mkdir(exportDir, destDir);
          fs.writeFileSync(`${exportDir}/${destDir}/${pageCfg.fileName}`, res);
          this.progress.finished++;
        }
      }
      this.loading = false;
    },
    convertPage({ dir, pageCfg, row }) {
      return new Promise((resolve, reject) => {
        const { width, height, fileName, publicStyle } = pageCfg;
        let canvas = document.getElementById("canvas");
        let context = canvas.getContext('2d');
        canvas.width = width;
        canvas.height = height;
        const elements = [];
        for(let placeholder of pageCfg.placeholders) {
          let { type, name, style, randomRange } = placeholder;
          if (publicStyle) {
            style = Object.assign({}, publicStyle, style);
          }
          if(!style) {
            continue;
          }
          switch(type) {
            case 'text':
              elements.push({
                html: row[name],
                style: this.getStyles(style),
                type
              });
            break;
            case 'year':
              elements.push({
                html: ((new Date).getFullYear()),
                style: this.getStyles(style),
                type
              });
            break;
            case 'day':
              elements.push({
                html: ((new Date).getDate()),
                style: this.getStyles(style),
                type
              });
            break;
            case 'month':
              elements.push({
                html: ((new Date).getMonth() + 1),
                style: this.getStyles(style),
                type
              });
            break;
            case '盖章':
            case '仿真盖章':
            case '签名':
              let companyName = '';
              if (placeholder.value) {
                companyName = placeholder.value.trim();
              } else if (row[name]) {
                companyName = row[name].trim();
              } else {
                continue;
              }
              if(type === '盖章' && !style.transform) {
                style.transform = `rotate(${(Math.random() * 360).toFixed(2)}deg)`;
              }
              if (type === '仿真盖章') {
                const [ x = 4, y = 4 ] = randomRange;
                style.transform = `rotate(${(Math.random() * 360).toFixed(2)}deg)`;
                style.marginTop = parseFloat(Math.random() * y) + '%';
                style.marginLeft = parseFloat(Math.random() * x) + '%';
                console.log(x, y);
              }
              const file = this.getPicFile(dir, companyName);
              elements.push({
                html: `<img style="max-width: 100%; max-height: 100%;" src="${file}" />`,
                style: this.getStyles(style),
                type
              });
            break;
          }
        }
        let html = `${htmlStyles}
        <div style="position:absolute; top:0; left:0; width: ${width}; height: ${height}">
          <img style="width: ${width}; height: ${height}" src="${dir}/${fileName}" />`;
        for(let item of elements) {
          html += `<div style="position:absolute; ${item.style}">${item.html}</div>`;
        }
        html += '</div>';
        rasterizeHTML.drawHTML(html, canvas)
        .then(res => {
          context.drawImage(res.image, width, height);
          canvas.toBlob(blob => {
            const reader = new FileReader();
            reader.addEventListener("loadend", function() {
               resolve(Buffer.from(reader.result));
            });
            reader.readAsArrayBuffer(blob);
          }, 'image/jpeg', 95);
        });
      });
    },
    selectFile() {
      const files = dialog.showOpenDialog({
        title: '选择 XLSX 文件',
        filters: [
          { name: 'xlsx文件', extensions: [ 'xlsx'] }
        ]
      });
      if(files) {
        if(this.filePath) {
          this.unwatch(this.filePath);
        }
        this.filePath = files[0];
        this.parseXLSX(this.filePath);
        this.watch(this.filePath);
      }
    },
    watch(file) {
      this.watcher = chokidar.watch(file, {}).on('all', (event, path) => {
        this.parseXLSX(this.filePath);
      });
    },
    unwatch(file) {return
      console.log(this.watcher)
      if(this.watcher) {
        this.watcher.unwatch(file);
        this.watcher.close();
      }
    },
    selectDir(type) {
      let defaultPath = '';
      if(this[type]) {
        defaultPath = this[type];
      }
      const dirs = dialog.showOpenDialog({
        title: '选择目录',
        properties: [ 'openDirectory' ],
        defaultPath
      });
      if(dirs) {
        this[type] = dirs[0];
      }
    },
    openurl() {
      shell.openExternal('http://wx.shouwang.io/doc2xlsx/');
    },
    parseXLSX(file) {
      const workbook = XLSX.readFile(file);
      // workbook.SheetNames
      const [ theFirstSheetName ] = workbook.SheetNames;
      const theFirstSheet = workbook.Sheets[theFirstSheetName];
      const cellNames = Object.keys(theFirstSheet).filter( name => !name.startsWith('!') );
      const tableData = [];
      this.fields = [];
      for(let cellName of cellNames) {
        const [, col, row] = cellName.match(/([A-Z]+)([0-9]+)/);
        const index = parseInt(row);
        const cellIndex = index - 2;
        if (cellIndex === -1) {
          this.fields.push(theFirstSheet[cellName].v);
          this.fieldsPosition.push(col);
        } else {
          const labelIndex = this.fieldsPosition.indexOf(col);
          const label = this.fields[labelIndex];
          if(cellIndex >= tableData.length) {
            tableData[cellIndex] = {};
          }
          tableData[cellIndex][label] = theFirstSheet[cellName].v;
        }
      }
      this.tableData = tableData;
    },
    handleOpen(key) {
      if(this.currentPage === key) {
        return;
      } else {
        this.currentPage = key;
      }
    },
    getDirPrefix() {
      const today = new Date();
      return `电子文书：${today.getMonth()+1}月${today.getDate()}日`
    },
    // node did not support recursive mode under v10.12 
    mkdir(prefix, dir) {
      if(fs.existsSync(prefix + dir)) {
        return true;
      }
      try {
        return fs.mkdirSync(prefix + dir);
      } catch (e) {
        const paths = dir.split('/');
        let tempPath = '';
        for(let path of paths) {
          tempPath += ('/' + path);
          this.mkdir(prefix, tempPath);
        }
      }
    },
    createName(row) {
      if(row['原账号名称']) {
        return `${row.原账号名称}-${row.原账号原始ID}`;
      }
      if(row['原始ID']) {
        return `${row.账号名称}-${row.原始ID}`;
      }
      if(row['甲方']) {
        return `${row.甲方}`;
      }
      return '其他'
    },
    getStyles(styles) {
      const outputStyles = [];     
      for(let key in styles) {
        const styleKey = key.replace(/([A-Z]+)/g, (str)=> '-'+str.toLowerCase());
        const value = styles[key];
        outputStyles.push(`${styleKey}: ${value}`);
      }
      return outputStyles.join("; ");
    },
    saveCfg (key, value) {
      let cfg = window.localStorage.getItem('cfg');
      if (!cfg) {
        cfg = {};
      }
      cfg[key] = value;
      return window.localStorage.setItem('cfg', JSON.stringify(cfg));
    },
    readAllCfg () {
      let config = window.localStorage.getItem('cfg');
      return JSON.parse(config);
    },
    readDocCfg() {
      if (!this.stampDir) {
        return false;
      }
      const cfgFile = `${this.stampDir}/配置.js`;
      if(require.cache[cfgFile]) {
        delete require.cache[cfgFile];
      }
      if (fs.existsSync(cfgFile)) {
        const { cfg } = require(cfgFile);
        this.cfg = cfg;
        this.$set(this.cfg, 'types', cfg.types);
      } else {
        return false;
      }
    },
    loadAllCfg() {
      let cfg = this.readAllCfg();
      if(!cfg) {
        return;
      }
      for(let key of Object.keys(cfg)) {
        this.$set(this, key, cfg[key]);
      }
    },
    loadFile() {
      if(this.filePath && fs.existsSync(this.filePath)) {
        this.parseXLSX(this.filePath);
        this.watch(this.filePath);
      }
    },
    saveAllCfg() {
      const cfg = {
        filePath: this.filePath,
        currentPage: this.currentPage,
        exportType: this.exportType,
        exportDir: this.exportDir,
        stampDir: this.stampDir
      };
      return window.localStorage.setItem('cfg', JSON.stringify(cfg));
    },
    listenUnload() {      
      window.addEventListener('beforeunload', (e) => {
        this.saveAllCfg();
      });
    },
    openExportDir() {
      shell.openItem(`${this.exportDir}/${this.getDirPrefix()}`);
    }
  },
  async mounted() {
    this.loadAllCfg();
    this.loadFile();
    this.listenUnload();
    this.readDocCfg();
  }
})