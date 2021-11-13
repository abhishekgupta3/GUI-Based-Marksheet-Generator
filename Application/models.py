from django.db import models

class myuploadfile(models.Model):
    master_roll = models.FileField(upload_to="", null=True)
    responses = models.FileField(upload_to="", null=True)
    postive_marks = models.IntegerField(null=True)
    negative_marks = models.IntegerField(null=True)

    