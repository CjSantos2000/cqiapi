from django.db import models

# Create your models here.
class AutoFetchModel(models.Model):
    user_id = models.CharField(max_length=500, null=True, blank=True)
    name = models.TextField(max_length=500, null=True, blank=True)
    category = models.CharField(max_length=500, null=True, blank=True)
    def __str__(self):
        return str(self.name)


class DownloadModel(models.Model):
    user_id = models.CharField(max_length=500, null=True, blank=True)
    record_id = models.CharField(max_length=500, null=True, blank=True)
    form_name = models.CharField(max_length=500, null=True, blank=True)
    def __str__(self):
        return str(self.user_id)