# Generated by Django 5.0.1 on 2024-02-06 10:22

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('app', '0001_initial'),
    ]

    operations = [
        migrations.AddField(
            model_name='autofetchmodel',
            name='category',
            field=models.CharField(blank=True, max_length=500, null=True),
        ),
    ]