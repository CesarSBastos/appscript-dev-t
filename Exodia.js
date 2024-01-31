function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu').addItem('Run', 'functionUpdate').addToUi();
}

const sheet = SpreadsheetApp.getActiveSheet();
const range = sheet.getDataRange();
const values = range.getValues();


class Aluno {
  constructor(id, name, absences, p1, p2, p3) {
    this.id = id;
    this.name = name;
    this.absences = absences;
    this.p1 = p1;
    this.p2 = p2;
    this.p3 = p3;
  }
  calcAverage() {
    let average = (this.p1 + this.p2 + this.p3) / 3;
    return average
  }


  getSituation(absences, i) {
    const baseabsences = 60 * 0.25;
    let situationRange = sheet.getRange(i + 1, 7);

    if (absences > baseabsences) {
      situationRange.setValue('Reprovado por falta').setBackground("red");
      let newRange = sheet.getRange(i + 1, 8).setValue(0).setBackground("red");      
    }
    return (this.absences)
  }

  getFinalNote(i,aluno) {
    if (this.calcAverage() < 50) {
      let notesRange = sheet.getRange(i + 1, 7).setValue('Reprovado por nota').setBackground("red");
      let newRange = sheet.getRange(i + 1, 8).setValue(0).setBackground("red");
    } else if (this.calcAverage() >= 50 && this.calcAverage() < 70) {
      let notesRange = sheet.getRange(i + 1, 7).setValue('Exame Final').setBackground("yellow");
      let newRange = sheet.getRange(i + 1, 8).setValue(aluno.getNaf()).setBackground("yellow")
    } else if (this.calcAverage() >= 70) {
      let notesRange = sheet.getRange(i + 1, 7).setValue('Aprovado').setBackground("green");
      let newRange = sheet.getRange(i + 1, 8).setValue(0).setBackground("green");
    }
  }

  getNaf() {
    const naf = (50 * 2) - this.calcAverage()
    return parseInt(naf);
  }
}


function functionUpdate() {

  for (i = 3; i < values.length; i++) {
    let id = values[i][0];
    let name = values[i][1];
    let absences  = values[i][2]
    let p1 = parseInt(values[i][3]);
    let p2 = parseInt(values[i][4]);
    let p3 = parseInt(values[i][5]);

    let aluno = new Aluno(id, name, absences , p1, p2, p3)

    let finalNote = aluno.getFinalNote(i,aluno)
    let situationAlumn = aluno.getSituation(absences, i)

  }

}





