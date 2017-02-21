module.export = class OpenXMLRelationship {

  constructor (pkg, part, relId, relType, target, targetMode) {

    this.fromPkg = pkg
    this.fromPart = part
    this.relId = relId
    this.relType = relType
    this.target = target
    this.targetMode = targetMode

    if (!targetMode) this.targetMode = 'Internal'
    if (targetMode === 'External') {
      this.targetFullName = this.target
      return
    }

    let workingTarget = target
    let workingCurrentPath = '/'
    if (this.fromPart) {
      let slashIndex = this.fromPart.uri.lastIndexOf('/')
      if (slashIndex !== -1) workingCurrentPath = this.fromPart.uri.substring(0, slashIndex) + '/'
    }

    while (workingTarget.startWith('../')) {
      if (workingCurrentPath.endsWith('/')) {
        workingCurrentPath = workingCurrentPath.substring(0, workingCurrentPath.length - 1)
      }

      let indexOfLastSlash = workingCurrentPath.lastIndexOf('/')
      workingCurrentPath = workingCurrentPath.substring(0, indexOfLastSlash + 1)
      workingTarget = workingTarget.substring(3)
    }

    this.targetFullName = workingCurrentPath + workingTarget
  }
}
